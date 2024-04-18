//CHANGE THESE
var email = 'EMAIL'
var doc = DocumentApp.openById("GOOGLE_DOCS_ID") //This is the string of characters after the main link

var pt_to_cm = 28.3465
var body = doc.getBody()
var body_txt = body.editAsText()
var body_pars = body.getParagraphs()
var body_pars_txt = body.getText().split("\n")
var transtions = ["INTERCUT"]

function format_doc() {
  //reset_doc()
  var previous_par = body_pars_txt[0]
  var scene_counter = 0
  for (var i = 0; i < body_pars_txt.length; i = i + 1) {
    var par = body_pars[i]
    var par_txt = par.getText()
    //Ignore Title and Subtitle
    if (par.getHeading() == DocumentApp.ParagraphHeading.TITLE || par.getHeading() == DocumentApp.ParagraphHeading.SUBTITLE){
      Logger.log("title or subtitle")
    }
    //Empty Paragraphs
    else if (par_txt == "" || par_txt == " "){
      previous_par = "null"
    }
    //Format Transitions
    else if(par_txt.toUpperCase().slice(par_txt.length-3, par_txt.length) == "TO:" || par_txt.toUpperCase().slice(par_txt.length-3, par_txt.length) == "IN:" || par_txt.toUpperCase().slice(par_txt.length-4, par_txt.length) == "OUT:"){
      par.setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
    }
    //Format Parentheticals
    else if (par_txt[0] == "(" && par_txt[par_txt.length-1] == ")"){
      par.setIndentFirstLine(3.81*pt_to_cm)
      par.setIndentStart(3.81*pt_to_cm)
      par.setIndentEnd((17.78-2.54-9.84)*pt_to_cm)
    }
    //Format Dialogue
    else if(previous_par == previous_par.toUpperCase()){
      par.setIndentFirstLine(2.54*pt_to_cm)
      par.setIndentStart(2.54*pt_to_cm)
      par.setIndentEnd((17.78-2.54-10.85)*pt_to_cm)
    }
    //Ignore Format on Paragraphs with Periods 
    else if(par_txt[par_txt.length-1]=="."){
      var _ = 0
    }
    else if (par_txt == par_txt.toUpperCase()){
      //Format Transtions
      if (par_txt == transtions[0]){
        var _ = 0
      }
      //Format Sluglines
      else if(!isNaN(parseFloat(par_txt[0])) || par_txt.slice(0, 3) == "INT" || par_txt.slice(0, 3) == "EXT" || par_txt[0] == "#"){
        scene_counter+=1
        if(par_txt.slice(0, 3) == "INT" || par_txt.slice(0, 3) == "EXT"){
          par.setText(scene_counter.toString() + "\t" + par.getText())
        }
        else if (par_txt[0]=="#"){
          par.replaceText("#", scene_counter.toString())
        }
        else{
          var par_split = par.getText().split("\t")
          par.setText(scene_counter.toString() + "\t" + par_split[1])
        }
        par.setHeading(DocumentApp.ParagraphHeading.NORMAL)
        par.setHeading(DocumentApp.ParagraphHeading.HEADING1)
        par.editAsText().setBold(true)
        par.editAsText().setFontSize(12)
        par.setIndentFirstLine(-1.27*pt_to_cm)
        previous_par = "heading"
      }
      //Format Speaker
      else{
        par.setHeading(DocumentApp.ParagraphHeading.NORMAL) 
        previous_par = par_txt
        par.setIndentStart(5.08*pt_to_cm)
        par.setIndentFirstLine(5.08*pt_to_cm)
      }
    }
    //Format Heading 2
    else if(par_txt.slice(0, 12) == "Music plays:" || par_txt.slice(0, 12) == "Music Plays:"){
      par.setHeading(DocumentApp.ParagraphHeading.NORMAL)
      par.setHeading(DocumentApp.ParagraphHeading.HEADING2)
      par.editAsText().setFontSize(12)
      par.editAsText().setUnderline(true)
      par.editAsText().setBold(true)
      previous_par = par_txt
    }
    //Logger.log(i.toString() + " - " + par.getText() + "           " + i.toString() + " - " + body_pars_txt[i])
  }
}

function reset_doc(){
  body.setMarginLeft(3.81*pt_to_cm)
  body.setMarginRight(2.54*pt_to_cm)
  body.setMarginBottom(2.54*pt_to_cm)
  body.setMarginTop(2.54*pt_to_cm)
  for (var i = 0; i < body_pars.length; i = i + 1) {
    par = body_pars[i]
    par.setIndentFirstLine(0)
    par.setIndentStart(0)
    par.setIndentEnd(0)
  }
}

function email_error(rowNumber) {
  var sheet = SpreadsheetApp.getActiveSheet()
  var data = sheet.getDataRange().getValues()
  var rowData = data[rowNumber-1].join(" ")
  MailApp.sendEmail(email,
                    'Data in row ' + rowNumber,
                    rowData)
}

function report(message) {
  Logger.log(message)
  //MailApp.sendEmail(email, "Google Script Report", message)
}
