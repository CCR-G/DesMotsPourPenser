var myDoc = app.activeDocument;

//ASKING THE USER TO GIVE A LIST OF WORDS CORRESPONDING TO FILES

var userInput = prompt(
  "Write in the names of files without the extension",
  "word1, word2, etc…",
  "Names of text files"
);

//- Checking if the user canceled the script
if (userInput == null) {
  $.writeln("prompt canceled");
  exit();
}
$.writeln(userInput);

//SPLITING THE STRING INPUT INTO AN ARRAY AND LOOPING THROUGH IT

var myWordsArray = userInput.split(", ");
for (var i = 0; i < myWordsArray.length; i++) {
  //CREATING A URI WITH USER INPUT AND READING THE CORRESPONDING FILE

  var myWord = myWordsArray[i].toLowerCase();
  var txtFile = File(
    "C:\\Users\\Camilles\\Documents\\DesMotsPourPenser\\texte\\" +
      myWord +
      ".txt"
  );
  txtFile.encoding = "UTF8";
  txtFile.open("r");
  var fileContentsString = txtFile.read();
  txtFile.close();

  //CHECKING WHETHER THE FILE OF THE INPUT WORD IS EMPTY

  //- The file is not empty
  if (fileContentsString.length !== 0) {
    //- Checking whether the image for the word already exists (We use search & find hidden text we add for each word image)
    app.findChangeTextOptions = null;
    app.findTextPreferences = app.changeTextPreferences = null;
    app.findTextPreferences.findWhat = myWord;
    app.findTextPreferences.appliedParagraphStyle = "No-Print";
    var myFoundText = myDoc.findText();
    if (myFoundText.length !== 0) {
      buttons(); //- A word image already exists, we give a choice to the user
    } else {
      creerMot(); //- No word image by that name exists, we create one
    }
  }
  //- The file is empty
  else {
    alert('No file corresponding to the word "' + myWord + '"');
  }
}

//FUNCTION FOR DISPLAYING CHOICES BUTTONS
function buttons() {
  var w = new Window("dialog", 'The word "' + myWord + '" already exists…'),
    u,
    modifierBtn = w.add(
      "button",
      u,
      "Delete the existing image and create a new one"
    ); //- Choosing to delete the word image and create a new one
  (passerBtn = w.add("button", u, "Keep the existing image and move on")),
    (passerBtn.code = 0); //- Choosing to keep the existing word and move on to the next one
  modifierBtn.code = 1;

  w.preferredSize.width = 400;
  w.alignChildren = ["fill", "top"];

  passerBtn.onClick = modifierBtn.onClick = function () {
    w.close(this.code);
  };

  var result = w.show();

  if (result == 1) {
    myFoundText[0].parentTextFrames[0].parentPage.remove(); //- Deleting the page on which the word has been found
    myDoc.colors.itemByName(myWord).remove(); //- Deleting the previous word color
    creerMot();
  }
}

//FUNCTION FOR CREATING A NEW WORD IMAGE
function creerMot() {
  //ADDING A PAGE AT THE BEGINNING OF THE DOCUMENT
  myDoc.pages.add(LocationOptions.BEFORE, myDoc.pages[0]);

  //CREATING RANDOM COLOR

  var l = Math.round(Math.random() * 40 + 50); //Setting a number between 50 and 90
  var a = Math.round(Math.random() * 255 - 128);
  var b = Math.round(Math.random() * 255 - 128);

  var myRandomColor = myDoc.colors.add({
    name: myWord,
    model: ColorModel.process,
    space: ColorSpace.LAB,
  });

  myRandomColor.colorValue = [l, a, b];

  //CREATING BACKGROUND

  var myRectangles = myDoc.pages[0].rectangles;

  //- Page background
  var myPageBackground = myRectangles.add({
    geometricBounds: [0, 0, 540, 960],
    contentType: ContentType.UNASSIGNED,
    fillColor: myDoc.colors.item("myMainColor"),
    strokeWeight: 0,
  });

  //- Decoration background rectangles using random color
  var myLeftRectangle = myRectangles.add({
    geometricBounds: [0, 0, 405, 300],
    contentType: ContentType.UNASSIGNED,
    fillColor: myRandomColor,
    strokeWeight: 0,
  });
  var myRightRectangle = myRectangles.add({
    geometricBounds: [404, 824, 540, 960],
    contentType: ContentType.UNASSIGNED,
    fillColor: myRandomColor,
    strokeWeight: 0,
  });

  //- Text background with rounded corners
  var myTextBackgroundRadius = "26px";
  var myTextBackground = myRectangles.add({
    geometricBounds: [17.649, 21, 524, 942],
    contentType: ContentType.UNASSIGNED,
    fillColor: myDoc.colors.item("myTextBackground"),
    strokeWeight: 0,

    topLeftCornerOption: CornerOptions.ROUNDED_CORNER,
    topRightCornerOption: CornerOptions.ROUNDED_CORNER,
    bottomLeftCornerOption: CornerOptions.ROUNDED_CORNER,
    bottomRightCornerOption: CornerOptions.ROUNDED_CORNER,
    bottomLeftCornerRadius: myTextBackgroundRadius,
    bottomRightCornerRadius: myTextBackgroundRadius,
    topLeftCornerRadius: myTextBackgroundRadius,
    topRightCornerRadius: myTextBackgroundRadius,
  });

  //CREATING AND EDITING THE TEXT CONTENT

  //- Adding a text frame to the newly created page and adding file content as text content of the frame
  var myTextFrame = myDoc.pages[0].textFrames.add({
    geometricBounds: [173, 42.627, 472, 917.373],
  });
  myTextFrame.contents = fileContentsString;
  myTextFrame.contents += " \r\r";

  //- Deleting "[Mot du Jour] #" at the begining of the text content
  app.findTextPreferences = app.changeTextPreferences = null;
  app.findChangeTextOptions = null;
  app.findTextPreferences.findWhat = "[Mot du Jour] #";
  app.changeTextPreferences.changeTo = "";
  myDoc.changeText();

  //- Replacing hard returns to new paragraphs
  app.findTextPreferences.findWhat = ",^n";
  app.changeTextPreferences.changeTo = "^p";
  myDoc.changeText();

  //- Deleting the unneeded end of text content
  app.findChangeGrepOptions = null;
  app.findGrepPreferences = null;
  app.findGrepPreferences.findWhat = "\\n.+";
  app.changeGrepPreferences.changeTo = "";
  myDoc.changeGrep();

  //- Applying paragraph styles to text frame content
  myTextFrame.parentStory.paragraphs[0].appliedParagraphStyle = "Word";
  myTextFrame.parentStory.paragraphs[1].appliedParagraphStyle = "Definition";

  //CREATING EXTRA DECORATION BLOCKS

  //- Twitter account name
  var myTitle = myDoc.pages[0].textFrames.add({
    geometricBounds: [34.751, 544, 52.175, 917.619],
  });
  myTitle.contents = "Des mots pour penser";
  myTitle.parentStory.paragraphs[0].appliedParagraphStyle =
    "Des mots pour penser";

  //- Little cube at the bottom of the page
  var myRightRectangle = myRectangles.add({
    geometricBounds: [482.155, 33.963, 507.41, 59.218],
    contentType: ContentType.UNASSIGNED,
    fillColor: myDoc.colors.item("myMainColor"),
    strokeWeight: 0,
    bottomLeftCornerOption: CornerOptions.ROUNDED_CORNER,
    bottomLeftCornerRadius: "10px",
  });

  //- Anchored hashtag
  var myHashtag = myDoc.pages[0].textFrames.add();
  myHashtag.contents = "#MotDuJour";
  myHashtag.anchoredObjectSettings.insertAnchoredObject(
    myTextFrame.parentStory.paragraphs[0].insertionPoints[-2],
    AnchorPosition.ANCHORED
  );
  myHashtag.applyObjectStyle(myDoc.objectStyles.itemByName("Hashtag"));
  myHashtag.parentStory.paragraphs[0].fillColor = myRandomColor;

  //- Non-printable text frame containing the name of the file used (for search purpose)
  var myFileNameTextFrame = myDoc.pages[0].textFrames.add({
    geometricBounds: [34, 36, 52, 127],
  });
  myFileNameTextFrame.contents = myWord + ".txt";
  myFileNameTextFrame.nonprinting = true;
  myFileNameTextFrame.parentStory.paragraphs[0].appliedParagraphStyle =
    "No-Print";

  //EXPORTING

  //- Setting options for JPG export
  var myJpegPref = app.jpegExportPreferences;
  myJpegPref.jpegQuality = JPEGOptionsQuality.high; // low medium high maximum
  myJpegPref.exportResolution = 72;
  myJpegPref.jpegExportRange = ExportRangeOrAllPages.EXPORT_RANGE;
  myJpegPref.jpegRenderingStyle = JPEGOptionsFormat.BASELINE_ENCODING;
  myJpegPref.jpegColorSpace = JpegColorSpaceEnum.RGB;
  myJpegPref.embedColorProfile = false;
  myJpegPref.antiAlias = true;
  myJpegPref.pageString = myDoc.pages[0].name;

  //- Actually exporting the file to given folder
  myDoc.exportFile(
    ExportFormat.JPG,
    new File(
      "C:\\Users\\Camilles\\Documents\\DesMotsPourPenser\\illustrations\\" +
        myWord +
        ".jpg"
    ),
    false
  );
}

/* IDEAS FOR IMPROVING THE SCRIPT

V Properly commenting and translating code to English
- Reducing and cleaning code
- Replacing prompt to a nice window
- Prompt for selecting files or a folder to use for finding words
- Adding colors to a color folder
- Having one of the buttons preselected
- Adding the choice button "Do for all words"
- Catch errors everywhere

*/

/* THANK YOU TO :

- Colors : https://github.com/fabianmoronzirfas/extendscript/wiki/Colors-And-Swatches
- Button function : https://stackoverflow.com/questions/44119409/capture-users-click-indesign-scripting

*/
