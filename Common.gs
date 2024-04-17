/**
 * Callback for rendering the homepage card.
 * @return {CardService.Card} The card to show to the user.
 */
function onHomepage(e) {
  console.log(e);
  return createDebateCard();
}

// Create the UI
function createDebateCard() {
  // Code for creating buttons
  var actionCondense = CardService.newAction()
    .setFunctionName('condenseText')
    .setParameters({});
  var condenseButton = CardService.newTextButton()
    .setText("Condense")
    .setOnClickAction(actionCondense)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  // Code for creating buttons
  var pocketAction = CardService.newAction()
    .setFunctionName('headerFormat')
    .setParameters({ h: '1' });
  var pocketButton = CardService.newTextButton()
    .setText("Pocket")
    .setOnClickAction(pocketAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  // Code for creating buttons
  var hatAction = CardService.newAction()
    .setFunctionName('headerFormat')
    .setParameters({ h: '2' });
  var hatButton = CardService.newTextButton()
    .setText("Hat")
    .setOnClickAction(hatAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  // Code for creating buttons
  var blockAction = CardService.newAction()
    .setFunctionName('headerFormat')
    .setParameters({ h: '3' });
  var blockButton = CardService.newTextButton()
    .setText("Block")
    .setOnClickAction(blockAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  // Code for creating buttons
  var tagAction = CardService.newAction()
    .setFunctionName('headerFormat')
    .setParameters({ h: '4' });
  var tagButton = CardService.newTextButton()
    .setText("Tag")
    .setOnClickAction(tagAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  // Code for creating buttons
  var citationAction = CardService.newAction()
    .setFunctionName('customStyling')
    .setParameters({ opt: '0' });
  var citationButton = CardService.newTextButton()
    .setText("Citation")
    .setOnClickAction(citationAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  // Code for creating buttons
  var underlineAction = CardService.newAction()
    .setFunctionName('customStyling')
    .setParameters({ opt: '1' });
  var underlineButton = CardService.newTextButton()
    .setText("Underline")
    .setOnClickAction(underlineAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  // Code for creating buttons
  var emphasisAction = CardService.newAction()
    .setFunctionName('customStyling')
    .setParameters({ opt: '2' });
  var emphasisButton = CardService.newTextButton()
    .setText("Emphasis")
    .setOnClickAction(emphasisAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  // Code for creating buttons
  var highlightAction = CardService.newAction()
    .setFunctionName('customStyling')
    .setParameters({ opt: '3' });
  var highlightButton = CardService.newTextButton()
    .setText("Highlight")
    .setOnClickAction(highlightAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  // Code for creating buttons
  var clearAction = CardService.newAction()
    .setFunctionName('customStyling')
    .setParameters({ opt: '4' });
  var clearButton = CardService.newTextButton()
    .setText("Clear")
    .setOnClickAction(clearAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);


  // Code for pasting without formatting
  var pasteNoFormattingAction = CardService.newAction()
    .setFunctionName('verbatimPaste');
  var pasteButton = CardService.newTextButton()
    .setText("Paste")
    .setOnClickAction(pasteNoFormattingAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  // TODO TIMER...

  // Code for creating buttons
  // var uploadAction = CardService.newAction()
  //   .setFunctionName('openPublicForum')
  //   .setParameters({});
  // var uploadButton = CardService.newTextButton()
  //   .setText("Upload")
  //   .setOnClickAction(uploadAction)
  //   .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  // Add all the buttons to a single set
  var buttonSet1 = CardService.newButtonSet()
    .addButton(condenseButton);
  // .addButton(uploadButton);

  var buttonSet2 = CardService.newButtonSet()
    .addButton(pocketButton)
    .addButton(hatButton)
    .addButton(blockButton)
    .addButton(tagButton);
  var buttonSet3 = CardService.newButtonSet()
    .addButton(citationButton)
    .addButton(underlineButton)
    .addButton(emphasisButton)
    .addButton(highlightButton)
    .addButton(clearButton)
    .addButton(pasteButton);

  // Assemble the cards
  var section1 = CardService.newCardSection()
    .setHeader("Research")
    .addWidget(buttonSet1);
  var section2 = CardService.newCardSection()
    .setHeader("Organize")
    .addWidget(buttonSet2);
  var section3 = CardService.newCardSection()
    .setHeader("Format")
    .addWidget(buttonSet3);
  var card = CardService.newCardBuilder()
    .addSection(section1)
    .addSection(section2)
    .addSection(section3);
  return card.build();

}

// Sample action for unfinished functions...
function onClickOne(e) {
  // // Create a new card with the same text.
  var card = createDebateCard()

  // Create an action response that instructs the add-on to replace
  // the current card with the new one.
  var navigation = CardService.newNavigation()
    .updateCard(card);
  var actionResponse = CardService.newActionResponseBuilder()
    .setNavigation(navigation);
  return actionResponse.build();
  // console.log(e);
  // return;
}

function condenseText() {
  var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();

  if (selection) {
    var elements = selection.getRangeElements();
    var condensedText = '';

    for (var i = 0; i < elements.length; i++) {
      var element = elements[i].getElement();
      var text = '';

      // Check if the element is a text element and extract text
      if (element.editAsText) {
        var textElement = element.editAsText();

        // Determine the text range to extract
        var startIndex = elements[i].isPartial() ? elements[i].getStartOffset() : 0;
        var endIndex = elements[i].isPartial() ? elements[i].getEndOffsetInclusive() : textElement.getText().length - 1;

        text = textElement.getText().substring(startIndex, endIndex + 1);
      }

      // Append text to condensedText, adding a space if needed
      if (condensedText && text) {
        condensedText += ' ';
      }
      condensedText += text;
    }

    // Replace multiple newline characters and trailing spaces with a single space, then trim
    condensedText = condensedText.replace(/(\n+\s*)+/g, ' ').trim();

    if (condensedText) {
      // Insert the condensed text at the start of the selection and delete the original content
      var firstElement = elements[0].getElement();
      var parent = firstElement.getParent();
      var childIndex = parent.getChildIndex(firstElement);

      parent.insertParagraph(childIndex, condensedText);

      // Delete the original selected content
      for (var i = 0; i < elements.length; i++) {
        var element = elements[i].getElement();
        var parentElement = element.getParent();
        var elementType = element.getType();

        if (elementType === DocumentApp.ElementType.PARAGRAPH) {
          parentElement.removeChild(element);
        } else if (element.editAsText) {
          var textElement = element.editAsText();
          textElement.setText('');
        }
      }
    }
  } else {
    DocumentApp.getUi().alert('Please select some text in the document.');
  }
}



// DONE!
function headerFormat(e) {
  var h = e.parameters.h;
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var selection = doc.getSelection();

  if (selection) {
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];

      if (element.isPartial()) {
        var textElement = element.getElement().asText();
        var startOffset = element.getStartOffset();
        var endOffset = element.getEndOffsetInclusive();

        var beforeText = textElement.getText().substring(0, startOffset);
        var selectedText = textElement.getText().substring(startOffset, endOffset + 1);
        var afterText = textElement.getText().substring(endOffset + 1);

        var parentElement = textElement.getParent();
        var index = body.getChildIndex(parentElement);

        // Replace the original text element with the before text, selected text, and after text as separate paragraphs
        textElement.removeFromParent();

        if (afterText) {
          body.insertParagraph(index + 1, afterText);
        }
        var selectedParagraph = body.insertParagraph(index, selectedText);
        if (h === '1') {
          selectedParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);
        }
        else if (h === '2') {
          selectedParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING2);
        }
        else if (h === '3') {
          selectedParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING3);
        }
        else if (h === '4') {
          selectedParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING4);
        }
        // selectedParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);
        if (beforeText) {
          body.insertParagraph(index, beforeText);
        }
      } else {
        // For whole selected elements, no action needed as it's already a separate paragraph
        DocumentApp.getUi().alert('The selected text is already a separate paragraph.');
      }
    }
  } else {
    DocumentApp.getUi().alert('Please select some text.');
  }
}

// Think this is done? Maybe add remove formatting from h1 etc.
function customStyling(e) {
  var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();

  if (selection) {
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      var textElement = element.getElement();

      // Check if the element supports text editing (skip empty lines and other non-text elements)
      if (textElement.editAsText) {
        var text = textElement.editAsText();

        // Determine the correct start and end offsets
        var startOffset = element.isPartial() ? element.getStartOffset() : 0;
        var endOffset = element.isPartial() ? element.getEndOffsetInclusive() : text.getText().length - 1;

        // Check if the element is not just an empty line or space
        if (text.getText().trim().length > 0) {
          // Remove existing formatting
          text.setBold(startOffset, endOffset, false);
          text.setItalic(startOffset, endOffset, false);
          text.setUnderline(startOffset, endOffset, false);
          text.setForegroundColor(startOffset, endOffset, "#000000");
          text.setBackgroundColor(startOffset, endOffset, "#FFFFFF");
          text.setFontSize(startOffset, endOffset, null);

          // Apply new styles based on the styleOption
          switch (e.parameters.opt) {
            case '0':
              // Set font to Arial, size to 12, and color to grey
              text.setFontFamily(startOffset, endOffset, "Arial");
              text.setFontSize(startOffset, endOffset, 12);
              text.setForegroundColor(startOffset, endOffset, "#989898"); // Grey (152, 152, 152)
              break;
            case '1':
              // Underline the text
              text.setUnderline(startOffset, endOffset, true);
              break;
            case '2':
              // Bold and underline the text
              text.setUnderline(startOffset, endOffset, true);
              text.setBold(startOffset, endOffset, true);
              text.setUnderline(startOffset, endOffset, true);
              break;
            case '3':
              // Bold, underline, and highlight the text in yellow
              text.setBold(startOffset, endOffset, true);
              text.setUnderline(startOffset, endOffset, true);
              text.setBackgroundColor(startOffset, endOffset, "#FFFF00"); // Yellow
              break;
            case '4':
              // Do nothing
              break;
            default:
              Logger.log('Invalid option. Please choose 1, 2, 3, or 4.');
          }
        }
      }
    }
  } else {
    Logger.log('No text selected.');
  }
}

function openPublicForum() {
  var html = '<html><body>';
  html += '<iframe src="https://opencaselist.com/hspf23" style="width:800px; height:500px;"></iframe>';
  html += '</body></html>';
  var ui = HtmlService.createHtmlOutput(html)
    .setWidth(820) // Adding a bit more to account for potential scrollbars
    .setHeight(520);
  DocumentApp.getUi().showModalDialog(ui, 'OpenCaseList Page');
}

function verbatimPaste() {
  var ui = DocumentApp.getUi();
  var result = ui.prompt(
    'Paste Clipboard Text',
    'Please paste the text from your clipboard:',
    ui.ButtonSet.OK_CANCEL);

  var button = result.getSelectedButton();
  var clipboardData = result.getResponseText();

  if (button == ui.Button.OK) {
    var doc = DocumentApp.getActiveDocument();
    var selection = doc.getSelection();

    // Updated fix:
    if (selection) {
      var elements = selection.getRangeElements();

      // We'll assume you want to replace the first text element  
      if (elements.length > 0 && elements[0].getElement().getType() == DocumentApp.ElementType.TEXT) {
        var textElement = elements[0].getElement().asText();
        textElement.setText(clipboardData);
      } else {
        // No text element selected, insert at cursor
        var cursor = doc.getCursor();
        cursor.insertText(clipboardData);
      }
    } else {
      // No selection, insert at cursor
      var cursor = doc.getCursor();
      cursor.insertText(clipboardData);
    }
  }
}
