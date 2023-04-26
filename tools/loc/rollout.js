import parseMarkdown from './helix/parseMarkdown.bundle.js';
import { mdast2docx } from './helix/mdast2docx.bundle.js';
import { docx2md } from './helix/docx2md.bundle.js';

import {
  connect as connectToSP,
  copyFileAndUpdateMetadata,
  getFileMetadata,
  getFileVersionInfo,
  getVersionOfFile,
  saveFileAndUpdateMetadata, updateFile,
} from './sharepoint.js';
import { getUrlInfo, loadingON, simulatePreview, stripExtension } from './utils.js';

const types = new Set();
const hashToContentMap = new Map();

function getBlockName(node) {
  if (node.type === 'gridTable') {
    let blockNameNode = node;
    try {
      while (blockNameNode.type !== 'text' && blockNameNode?.children.length > 0) {
        // eslint-disable-next-line prefer-destructuring
        blockNameNode = blockNameNode.children[0];
      }
      const fullyQualifiedBlockName = blockNameNode.value;
      if (fullyQualifiedBlockName) {
        const blockClass = fullyQualifiedBlockName.indexOf('(');
        let blockName = fullyQualifiedBlockName;
        if (blockClass !== -1) {
          blockName = fullyQualifiedBlockName.substring(0, blockClass).trim();
        }
        return blockName.toLowerCase().trim();
      }
    } catch (error) {
      // eslint-disable-next-line no-console
      console.log('Unable to get component name', error);
    }
  }
  return undefined;
}

function getNodeType(node) {
  return getBlockName(node) || node.type;
}

// Extracts and hashes the data in the mdast
// Generates 3 data structures
// 1. hashToIndex -> (Map) -> hash value to index (index defines the location/order of the content)
// 2. arrayWithTypeAndHash -> (Array) -> stores content type and equivalent hash value of the content in each array index
// 3. hashToContentMap -> (Map) -> stores hash value to node (node constains the mdast structure for heading, para etc.)
function processMdast(nodes) {
  // hash value to index (index defines the location/order of the content)
  const hashToIndex = new Map();
  // stores content type and equivalent hash value of the content in each array index
  const arrayWithTypeAndHash = [];
  let index = 0;
  // structure is a string holding the content structure of the entire doc. Eg: -heading-para-text-howto-para etc.
  let structure = '';
  nodes.forEach((node) => {
    // type of the node. eg: heading, text, para etc.
    const nodeType = getNodeType(node);
    types.add(nodeType);
    structure = `${structure}-${nodeType}`;
    // hash value of the node being processed
    const hash = objectHash.sha1(node);

    arrayWithTypeAndHash.push({ type: nodeType, hash });
    hashToIndex.set(hash, index);
    hashToContentMap.set(hash, node);
    index += 1;
  });

  // returns the generated map, array and structure data of the entire document
  return { hashToIndex, arrayWithTypeAndHash, structure };
}

async function getMd(path) {
  const mdPath = `${path}.md`;
  await simulatePreview(mdPath);
  const mdFile = await fetch(`${getUrlInfo().origin}${mdPath}`);
  return mdFile.text();
}

function getMdastFromMd(mdContent) {
  const state = { content: { data: mdContent }, log: '' };
  parseMarkdown(state);
  return state.content.mdast;
}

async function getMdast(path) {
  const mdContent = await getMd(path);
  return getMdastFromMd(mdContent);
}

// eslint-disable-next-line no-unused-vars
async function getProcessedMdastFromPath(path) {
  const mdast = await getMdast(path);
  const nodes = mdast.children || [];
  return processMdast(nodes);
}

async function getProcessedMdast(mdast) {
  const nodes = mdast.children || [];
  return processMdast(nodes);
}

async function persistDoc(srcPath, docx, dstPath) {
  try {
    await saveFileAndUpdateMetadata(srcPath, docx, dstPath, { RolloutStatus: 'Merged' });
  } catch (error) {
    // eslint-disable-next-line no-console
    console.log('Failed to save file', error);
  }
}

async function persist(srcPath, mdast, dstPath) {
  try {
    const docx = await mdast2docx(mdast);
    await persistDoc(srcPath, docx, dstPath);
  } catch (error) {
    // eslint-disable-next-line no-console
    console.log('Failed to save file', error);
  }
}

function updateChangesMap(changesMap, key, value) {
  if (!changesMap.has(key)) {
    changesMap.set(key, value);
  }
}

// left = processed data Map for the previously rolled out langstore file (langstoreBase)
// right
//    = processed data Map for the current langstore file that needs to be rolled out
//    OR
//    = processed data Map for the livecopy file
function getChanges(left, right) {
  const leftArray = left.arrayWithTypeAndHash;
  const leftHashToIndex = left.hashToIndex;
  const rightArray = right.arrayWithTypeAndHash;
  const rightHashToIndex = right.hashToIndex;
  const leftLimit = leftArray.length - 1;
  const rightLimit = rightArray.length - 1;

  // Holds the data (hash <-> change type) for the changed/unchanged content between langstoreBase and livecopy doc.
  const changesMap = new Map();
  // Holds content hash of livecopy block where the block type is the same between langstoreBase and livecopy at a
  // specific index but the content within the block was updated
  const editSet = new Set();

  let rightPointer = 0;
  for (let leftPointer = 0; leftPointer <= leftLimit; leftPointer += 1) {
    const leftElement = leftArray[leftPointer];
    if (leftPointer <= rightLimit) {
      rightPointer = leftPointer;
      const rightElement = rightArray[rightPointer];
      if (leftElement.hash === rightElement.hash) {
        // if there is not change between the blocks
        updateChangesMap(changesMap, leftElement.hash, { op: 'nochange' });
      } else if (rightHashToIndex.has(leftElement.hash)) {
        // if the block at this index in the livecopy/langstoreNow was moved to a different
        // location in the same document and the block also exists in the langstoreBase document
        // unedited, then mark the block in the livecopy/langstoreNow as 'added'
        updateChangesMap(changesMap, rightElement.hash, { op: 'added' });
      } else if (leftHashToIndex.has(rightElement.hash)) {
        // if the block at this index in the livecopy doc was deleted OR
        // in other words, langstoreBase has this block but liveccopy does not
        updateChangesMap(changesMap, leftElement.hash, { op: 'deleted' });
        updateChangesMap(changesMap, rightElement.hash, { op: 'added' });
      } else if (leftElement.type === rightElement.type) {
        // if both langstoreBase and livecopy doc has the same block type but different hash values.
        // this means that the block content was updated/edited. so this is marked as edited.
        updateChangesMap(changesMap, leftElement.hash, { op: 'edited', newHash: rightElement.hash });
        editSet.add(rightElement.hash);
      } else {
        // If the block in langstoreBase doc does not exist in livecopy doc or vice-versa.
        // TODO: Check if there are other conditions that apply here
        updateChangesMap(changesMap, leftElement.hash, { op: 'deleted' });
        updateChangesMap(changesMap, rightElement.hash, { op: 'added' });
      }
    } else {
      // if number of blocks in langstoreBase doc is more than the livecopy doc, then
      // everything extra in langstoreBase doc is marked as deleted
      updateChangesMap(changesMap, leftElement.hash, { op: 'deleted' });
    }
  }

  // if number of blocks in livecopy doc is more than the langstoreBase doc, then
  // everything extra in livecopy doc is marked as added (given that the condition below applies)
  for (let pointer = 0; pointer <= rightLimit; pointer += 1) {
    const rightElement = rightArray[pointer];
    // add the additional content to the changesMap only if
    // - it is not already part of the changesMap
    // - it is not already part of the the editSet (block that was edited at the same index location)
    if (!changesMap.has(rightElement.hash) && !editSet.has(rightElement.hash)) {
      changesMap.set(rightElement.hash, { op: 'added' });
    }
  }
  return changesMap;
}

/**
 * Get the content from the hash.
 */
function updateMergedMdast(mergedMdast, node, author) {
  const hash = node.opInfo.op === 'edited' ? node.opInfo.newHash : node.hash;
  const content = hashToContentMap.get(hash);
  if (node.opInfo.op !== 'edited' && node.opInfo.op !== 'nochange') {
    content.author = `${author} ${node.opInfo.op}`;
    content.action = node.opInfo.op;
  }
  mergedMdast.children.push(content);
}

// left = changes between previously rolled out langstore file (langstoreBase) and current langstore file
// right = changes between previously rolled out langstore file (langstoreBase) and the livecopy file
// Find differences and apply TrackChanges
function getMergedMdast(left, right) {
  const mergedMdast = { type: 'root', children: [] };
  let leftPointer = 0;
  // Holds LB<->LN changes
  const leftArray = Array.from(left.entries());

  function addTrackChangesInfo(author, action, root) {
    root.author = author;
    root.action = action;

    function addTrackChangesInfoToChildren(content) {
      if (content?.children) {
        const { children } = content;
        for (let i = 0; i < children.length; i += 1) {
          const child = children[i];
          if (child.type === 'text' || child.type === 'gtRow') {
            child.author = author;
            child.action = action;
          }
          if (child.type !== 'text') {
            addTrackChangesInfoToChildren(child);
          }
        }
      }
    }
    addTrackChangesInfoToChildren(root);
  }

  // Iterate over the LB<->LC (right) changeMap
  // eslint-disable-next-line no-restricted-syntax
  for (const [rightHash, rightOpInfo] of right) {
    if (left.has(rightHash)) {
      for (leftPointer; leftPointer < leftArray.length; leftPointer += 1) {
        const leftArrayElement = leftArray[leftPointer];
        const leftHash = leftArrayElement[0];
        const leftOpInfo = leftArrayElement[1];
        if (leftHash === rightHash) {
          const leftContent = hashToContentMap.get(leftOpInfo.op === 'edited' ? leftOpInfo.newHash : leftHash);
          if (rightOpInfo.op === 'nochange') {
            mergedMdast.children.push(leftContent);
          } else if (rightOpInfo.op === 'added') {
            addTrackChangesInfo('Langstore Version', 'added', leftContent);
            mergedMdast.children.push(leftContent);
          } else if (rightOpInfo.op === 'edited') {
            addTrackChangesInfo('Langstore Version', 'deleted', leftContent);
            mergedMdast.children.push(leftContent);
            const rightContent = hashToContentMap.get(rightOpInfo.newHash);
            addTrackChangesInfo('Regional Version', 'edited', rightContent);
            mergedMdast.children.push(rightContent);
          } else if (rightOpInfo.op === 'deleted') {
            addTrackChangesInfo('Langstore Version', 'deleted', leftContent);
            mergedMdast.children.push(leftContent);
          }
          leftPointer += 1;
          break;
        } else {
          updateMergedMdast(mergedMdast, { hash: leftHash, opInfo: leftOpInfo }, 'Langstore');
        }
      }
    } else {
      // These hashes in right do not exist in left. Mark these as edits from Region and merge.
      // updateMergedMdast(mergedMdast, { hash: rightHash, opInfo: rightOpInfo }, 'Region');
      const rightContent = hashToContentMap.get(rightHash);
      addTrackChangesInfo('Regional Version', 'edited', rightContent);
      mergedMdast.children.push(rightContent);
    }
  }
  // Check if there are excess elements in left array - if yes update them into the mergedMdast.
  if (leftPointer < leftArray.length) {
    // eslint-disable-next-line no-plusplus
    for (leftPointer; leftPointer < leftArray.length; leftPointer++) {
      const leftArrayElement = leftArray[leftPointer];
      // updateMergedMdast(
      //   mergedMdast,
      //   { hash: leftArrayElement[0], opInfo: leftArrayElement[1] },
      //   'Langstore',
      // );
      const leftContent = hashToContentMap.get(leftArrayElement[0]);
      addTrackChangesInfo('Langstore Version', 'deleted', leftContent);
      mergedMdast.children.push(leftContent);
    }
  }
  return mergedMdast;
}

// this logic is apparently from dexter
function noRegionalChanges(fileMetadata) {
  const lastModified = new Date(fileMetadata.Modified);
  const lastRollout = new Date(fileMetadata.Rollout);
  const diffBetweenRolloutAndModification = Math.abs(lastRollout - lastModified) / 1000;
  return false; //diffBetweenRolloutAndModification < 10;
}

async function safeGetVersionOfFile(filePath, version) {
  let versionFile;
  try {
    versionFile = await getVersionOfFile(filePath, version);
  } catch (error) {
    // Do nothing
  }
  return versionFile;
}

async function rollout(file, targetFolders, skipMerge = false) {
  // path to the doc that is being rolled out
  const filePath = file.path;
  await connectToSP();
  const filePathWithoutExtension = stripExtension(filePath);
  const fileName = filePath.substring(filePath.lastIndexOf('/') + 1);

  async function rolloutPerFolder(targetFolder) {
    // holds MD of the previous version of langstore file
    let langstorePrevMd;

    // live copy path of where the file is getting rolled out o
    const livecopyFilePath = `${targetFolder}/${fileName}`;

    // holds the status of whether rollout was successful
    const status = { success: true, path: livecopyFilePath };

    function udpateErrorStatus(errorMessage) {
      status.success = false;
      status.errorMsg = errorMessage;
      loadingON(errorMessage);
    }

    try {
      loadingON(`Rollout to ${livecopyFilePath} started`);
      // gets the metadata info of the live copy file
      const fileMetadata = await getFileMetadata(livecopyFilePath);
      // does the live copy file exist?
      const isfileNotFound = fileMetadata?.status === 404;
      // get langstore file prev version value from the rolled out live copy file's RolloutVersion value.
      // the RolloutVersion basically gives which version of langstore file was previously rolled out
      const langstorePrevVersion = fileMetadata.RolloutVersion;
      // get RolloutStatus value - skipped, merged etc.
      const previouslyMerged = fileMetadata.RolloutStatus;
      // get current version of the langstore file that will be rolled out
      const langstoreCurrentVersion = await getFileVersionInfo(filePath);

      // if regional file does not exist, just copy the langstore file to region
      if (isfileNotFound) {
        // just copy since regional document does not exist
        await copyFileAndUpdateMetadata(filePath, targetFolder);
        loadingON(`Rollout to ${livecopyFilePath} complete`);
        return status;
      }

      // if regional file exists but there are no changes in regional doc or if we are skipping merge altogether
      if (skipMerge || (noRegionalChanges(fileMetadata) && !previouslyMerged)) {
        await saveFileAndUpdateMetadata(
          filePath,
          file.blob,
          livecopyFilePath,
        );
        loadingON(`Rollout to ${livecopyFilePath} complete`);
        return status;
      }

      // if regional file exists AND there are changes in the regional file, then merge process kicks in.
      // ---
      if (!langstorePrevVersion) {
        // if for some reason the RolloutVersion metadata info is not available in region, rollout is NOT performed
        // Cannot merge since we don't have rollout version info.
        // eslint-disable-next-line no-console
        loadingON(`Cannot rollout to ${livecopyFilePath} since last rollout version info is unavailable`);
      } else if (langstoreCurrentVersion === langstorePrevVersion) {
        // If the same previosly rolled out file is being rolled out again, langstore file current version will be same as the_
        // _RolloutVersion (langstorePrevVersion) on the previously rolled out file in the respective region.
        // In this case, get the MD of the langstore file to be rolled out
        langstorePrevMd = await getMd(filePathWithoutExtension);
      } else {
        // If the langstore file current version is not same as the version of the rolled out file in the region, then
        // get the file (blob) of the previously rolled out version of the langstore file and then get the MD of that file

        // TODO: We can most probably get the mdast of latest langstore here (langstoreNow)
        // even the "else if" above may not be required

        const langstoreBaseFile = await safeGetVersionOfFile(filePath, langstorePrevVersion);
        if (langstoreBaseFile) {
          langstorePrevMd = await docx2md(langstoreBaseFile, { gridtables: true });
        }
      }

      // if the MD of the previously rolled out langstore file exists
      if (langstorePrevMd != null) {
        // get MDAST of the previously rolled out langstore file
        const langstorePrev = await getMdastFromMd(langstorePrevMd);
        // get processed data Map for the previously rolled out langstore file
        const langstoreBaseProcessed = await getProcessedMdast(langstorePrev);
        // get MDAST of the current langstore file that needs to be rolled out
        const langstoreNow = await getMdast(filePathWithoutExtension);
        // get processed data Map for the current langstore file that needs to be rolled out
        const langstoreNowProcessed = await getProcessedMdast(langstoreNow);

        const liveCopyPath = `${targetFolder}/${fileName}`;

        // if the content structure between the previously rolledout langstore file is NOT EQUAL TO current langstore file, then
        // skip rolling out the file
        // Set the RolloutStatus metadata on the livecopy file as 'Skipped'
        // TODO: handle merging content when current file structure in langstore (langstoreNow) is different than
        // the previously rolled out file structure (langstoreBase)
        // NOTE: commenting condition for testing regional edits
        // if (langstoreBaseProcessed.structure !== langstoreNowProcessed.structure) {
        //   await updateFile(liveCopyPath, null, { RolloutStatus: 'Skipped' });
        //   return status;
        // }

        // get MDAST of the livecopy file
        const livecopy = await getMdast(liveCopyPath.substring(0, liveCopyPath.lastIndexOf('.')));
        // get processed data Map for the livecopy file
        const livecopyProcessed = await getProcessedMdast(livecopy);

        // if the content structure between the previously rolledout langstore file is NOT EQUAL TO current livecopy file, then
        // skip rolling out the file
        // Set the RolloutStatus metadata on the livecopy file as 'Skipped'
        // TODO: handle merging content when previously rolled out file structure (langstoreBase) is different than the livecopy file structure
        // NOTE: commenting condition for testing regional edits
        // if (langstoreBaseProcessed.structure !== livecopyProcessed.structure) {
        //   await updateFile(liveCopyPath, null, { RolloutStatus: 'Skipped' });
        //   return status;
        // }

        // if the content structure is the same between the combinations checked above. i.e.
        // -- previously rolledout langstore file structure == current langstore file structure
        // AND
        // -- previously rolledout langstore file structure == current livecopy file structure
        // then get the changes between those combinations and generate the merged livecopy file
        // TODO: These should get called for structural changes (requires removing the if conditions above for structure check)
        const langstoreBaseToNowChanges = getChanges(langstoreBaseProcessed, langstoreNowProcessed);
        const langstoreBaseToRegionChanges = getChanges(langstoreBaseProcessed, livecopyProcessed);
        const livecopyMergedMdast = getMergedMdast(
          langstoreBaseToNowChanges,
          langstoreBaseToRegionChanges,
        );
        // save the merged livecopy file
        await persist(filePath, livecopyMergedMdast, livecopyFilePath);
        loadingON(`Rollout to ${livecopyFilePath} complete`);
      } else {
        udpateErrorStatus(`Rollout to ${livecopyFilePath} did not succeed. Missing langstore file`);
      }
    } catch (error) {
      udpateErrorStatus(`Rollout to ${livecopyFilePath} did not succeed. Error ${error.message}`);
    }
    return status;
  }

  const rolloutStatuses = await Promise.all(
    targetFolders.map((targetFolder) => rolloutPerFolder(targetFolder)),
  );
  return rolloutStatuses
    .filter(({ success }) => !success)
    .map(({ path }) => path);
}

export default rollout;
