// Format every link in your document using its own style

// Properties storage
const docProps = PropertiesService.getDocumentProperties();

const setDocProperties = (formObject) => {
  try {
    docProps.setProperties(formObject);
    Logger.log(`Saved properties: ${JSON.stringify(formObject)}`);
  } catch (err) {
    Logger.log(`Failed with error ${err.message}`);
  }
}

const getDocProperties = () => {
  try {
    const props = docProps.getProperties();
    Logger.log(`Requested properties: ${JSON.stringify(props)}`);
    return props;
  } catch (err) {
    Logger.log(`Failed with error ${err.message}`);
  }
}

const Style = docProps.getProperties();

// Add-on logic
const setStyle = (linkText, startRange, endRange) => {
  if (linkText && linkText.getType() === DocumentApp.ElementType.TEXT) {
    linkText.setForegroundColor(startRange, endRange, Style.color);
    linkText.setBold(startRange, endRange, (Style.bold === "true"));
    linkText.setItalic(startRange, endRange, (Style.italic === "true"));
    linkText.setUnderline(startRange, endRange, (Style.underline === "true"));
  }
}

const setLinksStyle = (element) => {
  let links = [];
  element = element || DocumentApp.getActiveDocument().getBody();

  if (element.getType() === DocumentApp.ElementType.TEXT) {
    const textObj = element.editAsText();
    const text = element.getText();

    let inUrl = false;
    let curUrl = {};

    for (let charIndex = 0; charIndex < text.length; charIndex++) {
      const url = textObj.getLinkUrl(charIndex);

      if (url === null || charIndex === text.length - 1) {
        if (inUrl) {
          if (url != null) curUrl.endOffsetInclusive = charIndex;
          curUrl.text = text.substring(curUrl.startOffset, curUrl.endOffsetInclusive + 1).trim();

          setStyle(textObj, curUrl.startOffset, curUrl.endOffsetInclusive);
          links.push(curUrl);

          inUrl = false;
          curUrl = {};
        }
      } else {
        if (!inUrl) {
          curUrl = {};

          curUrl.url = url;
          curUrl.startOffset = charIndex;

          inUrl = true;
        } else {
          curUrl.endOffsetInclusive = charIndex;
        }
      }
    }
    if (links.length > 0) Logger.log(links);
  } else {
    if (element.getType() != DocumentApp.ElementType.HORIZONTAL_RULE & element.getType() != DocumentApp.ElementType.DATE) {
      const numChildren = element.getNumChildren();
      for (let i = 0; i < numChildren; i++) {
        links = links.concat(setLinksStyle(element.getChild(i)));
      }
    }
  }
}

// UI part
const showModal = () => {
  const html = HtmlService.createTemplateFromFile('Config')
    .evaluate();

  const output = HtmlService.createHtmlOutput(html)
    .setWidth(220)
    .setHeight(170);

  DocumentApp.getUi()
    .showModalDialog(output, 'Set up your links');
}

const includeFile = (filename) => {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}

const onOpen = () => {
  DocumentApp.getUi()
    .createAddonMenu()
    .addItem('Style apply', 'setLinksStyle')
    .addSeparator()
    .addItem(`Style setup`, 'showModal')
    .addToUi();
}
