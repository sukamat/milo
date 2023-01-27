function getFloodgateUrl(url) {
  if (!url) {
    return undefined;
  }
  let urlArr = url.split('--');
  urlArr[1] += '-pink';
  return urlArr.join('--');
}

function getAnchorHtml(url, text) {
  return `<a href="${url}" target="_new">${text}</a>`;
}

function createColumn(innerHtml, classValue) {
  const tag = classValue === 'header' ? 'th' : 'td';
  const element = document.createElement(tag);
  if (innerHtml) {
    element.innerHTML = innerHtml;
  }
  return element;
}

export {
  getFloodgateUrl,
  getAnchorHtml,
  createColumn,
}
