import {
  createTag,
  getPathFromUrl,
} from '../../loc/utils.js';
import { getFloodgateUrl } from './utils.js';

function updateProjectInfo(project) {
  document.getElementById('project-url').innerHTML = `<a href='${project.sp}' title='${project.excelPath}'>${project.name}</a>`;
}

function getProjectDetailContainer() {
  const container = document.getElementsByClassName('project-detail')[0];
  container.innerHTML = '';
  return container;
}

function createRow(classValue = 'default') {
  return createTag('tr', { class: `${classValue}` });
}

function createColumn(innerHtml, classValue) {
  const tag = classValue === 'header' ? 'th' : 'td';
  const element = document.createElement(tag);
  if (innerHtml) {
    element.innerHTML = innerHtml;
  }
  return element;
}

function createHeaderColumn(innerHtml) {
  return createColumn(innerHtml, 'header');
}

function createTableWithHeaders() {
  const $table = createTag('table');
  const $tr = createRow('header');
  $tr.appendChild(createHeaderColumn('Source URL'));
  $tr.appendChild(createHeaderColumn('Source File'));
  $tr.appendChild(createHeaderColumn('Source File Info'));
  $tr.appendChild(createHeaderColumn('Floodgated URL'));
  $tr.appendChild(createHeaderColumn('Floodgated File'));
  $tr.appendChild(createHeaderColumn('Floodgated File Info'));
  $table.appendChild($tr);
  return $table;
}

function getAnchorHtml(url, text) {
  return `<a href="${url}" target="_new">${text}</a>`;
}

function getSharepointStatus(doc, isFloodgate) {
  let sharepointStatus = 'Connect to Sharepoint';
  let hasSourceFile = false;
  let modificationInfo = 'N/A';
  if (!isFloodgate && doc && doc.sp) {
    if (doc.sp.status === 200) {
      sharepointStatus = `${doc.filePath}`;
      hasSourceFile = true;
      modificationInfo = `By ${doc.sp?.lastModifiedBy?.user?.displayName} at ${doc.sp?.lastModifiedDateTime}`;
    } else {
      sharepointStatus = 'Source file not found!';
    }
  } else {
    if (doc.fg.sp.status === 200) {
      sharepointStatus = `${doc.filePath}`;
      hasSourceFile = true;
      modificationInfo = `By ${doc.fg.sp?.lastModifiedBy?.user?.displayName} at ${doc.fg.sp?.lastModifiedDateTime}`;
    } else {
      sharepointStatus = 'Floodgated file not found!';
    }
  }
  return { hasSourceFile, msg: sharepointStatus, modificationInfo };
}

function getLinkedPagePath(spShareUrl, pagePath) {
  return getAnchorHtml(spShareUrl.replace('<relativePath>', pagePath), pagePath);
}

function getLinkOrDisplayText(spViewUrl, docStatus) {
  const pathOrMsg = docStatus.msg;
  return docStatus.hasSourceFile ? getLinkedPagePath(spViewUrl, pathOrMsg) : pathOrMsg;
}

function showButtons(buttonIds) {
  buttonIds.forEach((buttonId) => {
    document.getElementById(buttonId).classList.remove('hidden');
  });
}

function updateProjectDetailsUI(projectDetail, config) {
  const container = getProjectDetailContainer();

  const $table = createTableWithHeaders();
  const spViewUrl = config.sp.shareUrl;
  const fgSpViewUrl = config.sp.fgShareUrl;

  projectDetail.urls.forEach((urlInfo, url) => {
    const $tr = createRow();
    const docPath = getPathFromUrl(url);

    // Source file data    
    const pageUrl = getAnchorHtml(url, docPath);
    $tr.appendChild(createColumn(pageUrl));
    const usEnDocStatus = getSharepointStatus(urlInfo.doc);
    const usEnDocDisplayText = getLinkOrDisplayText(spViewUrl, usEnDocStatus);
    $tr.appendChild(createColumn(usEnDocDisplayText));
    $tr.appendChild(createColumn(usEnDocStatus.modificationInfo));

    // Floodgated file data
    const fgPageUrl = getAnchorHtml(getFloodgateUrl(url), docPath);
    $tr.appendChild(createColumn(fgPageUrl));
    const fgDocStatus = getSharepointStatus(urlInfo.doc, true);
    const fgDocDisplayText = getLinkOrDisplayText(fgSpViewUrl, fgDocStatus);
    $tr.appendChild(createColumn(fgDocDisplayText));
    $tr.appendChild(createColumn(fgDocStatus.modificationInfo));

    $table.appendChild($tr);
  });

  container.appendChild($table);

  const showIds = ['startProject', 'copyFiles', 'promoteFiles', 'deleteFiles'];
  showButtons(showIds);
}

export {
  updateProjectInfo,
  updateProjectDetailsUI,
}
