// Globals
const sheet = SpreadsheetApp.getActiveSpreadsheet();
var actSheet = sheet.getActiveSheet();
const config_start_row = 1;
const CalendarSheetName = "Calendar";
const CalendarArtifcatName = "Calendar";
const CalendarTabArtifactName = "Calendar Tab Name";
const ConfigurationPage = "configuration";
const ArtifactsName = "Artifacts";
const ReleasesName = "Releases";
const ReleasesEndof = "End of Releases";
const NotesName = "Notes";
const NotesEndof = "End of Notes";
const ReleaseColorName = "Color";
const ReleaseMajorMarker = "Major";
const ReleaseBeginMarker = " ";
const ReleaseEndMarker = " ";
const SprintsName = "Sprints";
const SprintsEndof = "End of Sprints";
const DatabegingHere = "Calendar Data Begins Here"
const CalendarSheetStartColumn = 3;
const CalendarSheeetStartColumnOffet0 = CalendarSheetStartColumn - 1;
const CalendarSheetStartRow = 3;
const configurartionReleaseRow = 5;
const CalendarSprintDayMarker = "-"

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Calendar')
    .addItem('Update Release Calendar', 'main')
    .addItem('Update Column Widths', 'resetcolumns')
    .addToUi();
}
var readconfig = function () {
  //  const configurationpage = "configuration";
  SpreadsheetApp.setActiveSheet(sheet.getSheetByName(ConfigurationPage));
  actSheet = sheet.getActiveSheet();
  this.getData = function () {
    let lcol = actSheet.getDataRange().getLastColumn();
    let lrow = actSheet.getDataRange().getLastRow();
    let rdata = actSheet.getRange(1, 1, lrow, lcol).getValues();
    return rdata;
  }
  this.artifactlocationsSRN = function () {
    // find the location of artifact items - Sprints / Releases, Notes
    function locationrec() {
      this.type,
        this.begin,
        this.end
    }
    let rowz = this.getData().flatMap(x => [x[0]]);
    let dataloc = rowz.find(x => x == "Data");
    let data = rowz.slice(dataloc + 1);
    let result = [];
    let s = new locationrec();
    s.type = SprintsName;
    let r = new locationrec();
    r.type = ReleasesName;
    let n = new locationrec();
    n.type = NotesName;
    data.forEach((x, y) => {
      switch (x) {
        case SprintsName:
          s.begin = y;
          break;
        case SprintsEndof:
          s.end = y;
          break;
        case ReleasesName:
          r.begin = y;
          break;
        case ReleasesEndof:
          r.end = y;
          break;
        case NotesName:
          n.begin = y;
          break;
        case NotesEndof:
          n.end = y;
          break;
      }
    })
    result.push(s);
    result.push(r);
    result.push(n);
    return result;
  }
  this.getBackgroundsandFontColors = function () {
    const r = actSheet.getDataRange();
    const backgrounddata = r.getBackgrounds();
    const fontcolordata = r.getFontColors();;
    return { backgrounddata, fontcolordata };
  }

  this.getArtifacts = function (rdata) {
    function artifacts_rec() {
      this.artifact,
        this.title,
        this.configurationrowcolumn,
        this.displayrow
    }
    let result = [];
    rdata.forEach(x => {
      let a = new artifacts_rec();
      a.artifact = x[0];
      a.title = x[1];
      a.configurationrowcolumn = x[2];
      a.displayrow = x[3];
      result.push(a);
    })
    const afound = result.findIndex((e) => e.artifact == ArtifactsName);
    const rresults = result.slice(afound + 1);
    // Logger.log(afound);

    return rresults;
  }
  this.getReleases = function (locationsSRN, artifacts, rdata, rcolors) {
    function releaseconfig() {
      this.number,
        this.title,
        this.displayrow,
        this.displaycolor
    }
    function releasetypeRec() {
      this.major,
        this.colorcolumn,
        this.envs,
        this.displayrow,
        this.rowoffset,
        this.version,
        this.backgroundcolor,
        this.fontcolor,
        this.begin,
        this.end
    }
    function envrec() {
      this.majortype,
        this.major,
        this.majorcolor,
        this.majorfontcolor,
        this.env,
        this.envnumber,
        this.displayrow
    }
    const afound = artifacts.findIndex((e) => e.artifact == ReleasesName);
    const releaseName = artifacts[afound].artifact;
    let releaseDisplayRow = artifacts[afound].displayrow;
    let loc = locationsSRN.findIndex(x => x.type == ReleasesName);
    let rc = new Object;
    rc.row = locationsSRN[loc].begin + 2;
    rc.column = 1;
    let releases = [];
    //let rc = decode(artifacts[afound].configurationrowcolumn);
    let releaseType = rdata[rc.row - 1][rc.column - 1];
    let releaseMajor = rdata[rc.row - 1][rc.column];
    let rccolor = rdata[rc.row - 1];
    let rccolorIndex = rccolor.findIndex((c) => c == ReleaseColorName);
    let releaseMajorColor = rcolors.backgrounddata[rc.row - 1][rccolorIndex];
    let releaseMajorFontColor = rcolors.fontcolordata[rc.row - 1][rccolorIndex];
    //Logger.log("RC.ROW is %s",rc.row)
    let renvraw = JSON.parse(JSON.stringify(rdata[rc.row]));
    // Logger.log("length of 14 is %s",rdata[14].length);
    renvraw.shift();
    renvraw.shift();
    //     Logger.log("length of 14 is %s",rdata[14].length);
    let envs = [];
    renvraw.forEach((x) => {
      if (x !== "") {
        if (x !== ReleaseColorName) {
          let evrec = new envrec();
          evrec.majortype = releaseType;
          evrec.major = releaseMajor;
          evrec.majorcolor = releaseMajorColor;
          evrec.majorfontcolor = releaseMajorFontColor;
          evrec.displayrow = releaseDisplayRow;
          releaseDisplayRow++;
          evrec.env = x;
          evrec.envnumber = 0;
          envs.push(evrec);
          envs[envs.length - 1].envnumber = envs.length - 1; // set to the latest array element
        }
      }
    });
    rc.row++;
    //  envs.shift();
    for (let i = rc.row; i < rdata.length; i++) {
      // Logger.log('rdata [%s][%s] is %s', i, rc.column - 1, rdata[i][rc.column - 1])
      if (rdata[i][rc.column - 1] !== ReleasesEndof) {
        if (rdata[i][rc.column - 1] !== ReleaseMajorMarker) {
          let rcc = rc.column;
          envs.forEach((x) => {
            //   let paul = x;
            let r = new releasetypeRec();
            r.major = releaseMajor;
            r.envs = envs;
            r.version = rdata[i][rc.column];
            r.env = x.env;
            r.backgroundcolor = rcolors.backgrounddata[i][rccolorIndex];
            r.fontcolor = rcolors.fontcolordata[i][rccolorIndex];
            //Logger.log(" major %s, version %s, env %s, envs %s", r.major, r.version, r.env, r.envs);
            let efound = r.envs.findIndex((e) => e.env == r.env);
            r.rowoffset = r.envs[efound].envnumber;
            r.displayrow = x.displayrow;
            r.begin = rdata[i][rcc + 1];
            r.end = rdata[i][rcc + 2]
            if (r.end == "Color") {
              Logger.log("here")
            }
            if ((r.begin !== "") || (r.end !== "")) { //skip any releases missing dates
              releases.push(r);
            }
            rcc = rcc + 2;
          });
        }
        else {
          // new major release
          //  i++;
          releaseMajor = rdata[i][rc.column];
          rccolor = rdata[i];
          rccolorIndex = rccolor.findIndex((c) => c == ReleaseColorName);
          releaseMajorColor = rcolors.backgrounddata[i][rccolorIndex];
          releaseMajorFontColor = rcolors.fontcolordata[i][rccolorIndex];
          releaseDisplayRow++;
          renvraw = JSON.parse(JSON.stringify(rdata[i + 1]));
          //  renvraw = rdata[i + 1];
          renvraw.shift();
          renvraw.shift();
          envs = [];
          renvraw.forEach((x) => {
            if ((x !== "") && (x !== ReleaseColorName)) {
              //  if (x !== ReleaseColorName) {
              let evrec = new envrec();
              evrec.major = releaseMajor;
              evrec.majorcolor = releaseMajorColor;
              evrec.majorfontcolor = releaseMajorFontColor;
              evrec.displayrow = releaseDisplayRow;
              releaseDisplayRow++;
              evrec.env = x;
              envs.push(evrec);
              envs[envs.length - 1].envnumber = envs.length - 1; // set to the latest array element
              //   }
            }
          });
          //    envs.shift();
          i++;
        }
      }
      else {
        // Logger.log('rdata [%s][%s] is %s', i, rc.column - 1, rdata[i][rc.column - 1])
        i = rdata.length; // kill this loop 
        // Logger.log("End of releases!!!")
      }
    }
    return releases;

  }
  this.getSprints = function (locationsSRN, rdata, rcolors) {
    function sprintRec() {
      this.sprint,
        this.begin,
        this.end,
        this.backgroundcolor = "#ffffff", //"#ffffff"
        this.fontcolor = "#000000"
      this.markerbackgroundcolor = "#000000"
      this.markerfontcolor = "#ffffff"
    }
    // const afound = artifacts.findIndex((e) => e.artifact == SprintsName);
    // const sprintName = artifacts[afound].artifact;
    let sprints = [];
    let loc = locationsSRN.findIndex(x => x.type == SprintsName);
    //let rc = decode(artifacts[afound].configurationrowcolumn);
    let rc = new Object();
    rc.row = locationsSRN[loc].begin + 1;
    rc.column = 1;
    let sprintbackgroundcolor = rcolors.backgrounddata[rc.row - 1][3];
    let sprintfontcolor = rcolors.fontcolordata[rc.row - 1][3];
    let markerbackgroundcolor = rcolors.backgrounddata[rc.row - 1][4];
    let markerfontcolor = rcolors.fontcolordata[rc.row - 1][4];
    // rc.row++;
    for (let i = rc.row; i < rdata.length; i++) {
      if (rdata[i][rc.column - 1] !== SprintsEndof) {
        let s = new sprintRec();
        s.sprint = rdata[i][rc.column - 1];
        s.begin = rdata[i][rc.column];
        s.end = rdata[i][rc.column + 1];
        s.backgroundcolor = sprintbackgroundcolor;
        s.fontcolor = sprintfontcolor;
        s.markerbackgroundcolor = markerbackgroundcolor;
        s.markerfontcolor = markerfontcolor;
        sprints.push(s);
      }
      else {
        i = rdata.length;
      }
    }
    return sprints;
  }
  this.getNotes = function (locationsSRN, rdata, rcolors) {
    function noteRec() {
      this.Notetitle = "",
        this.NoteDate = "",
        this.NoteText = "",
        this.NoteBackgroundColor = "",
        this.NoteFontColor = ""
    }
    // const afound = artifacts[artifacts.findIndex((e) => e.artifact == NotesName)];
    // let rc = decode(afound.configurationrowcolumn);
    let loc = locationsSRN.findIndex(x => x.type == NotesName);
    let backgroundcolor = rcolors.backgrounddata[locationsSRN[loc].begin][0];
    let fontcolor = rcolors.fontcolordata[locationsSRN[loc].begin][0];
    let notes = [];
    let begin = locationsSRN[loc].begin;
    let end = locationsSRN[loc].end;
    for (let i = locationsSRN[loc].begin + 1; i <= locationsSRN[loc].end - 1; i++) {
      let nn = rdata[i];
      let n = new noteRec();
      n.Notetitle = nn[0];
      n.NoteDate = nn[1]
      n.NoteText = nn[2];
      n.NoteBackgroundColor = backgroundcolor;
      n.NoteFontColor = fontcolor;
      notes.push(n);
    }
    return notes;
  }
  this.verifyConfig = function () {
    /* lets verify the configuration as follows
    Artifacts exist
    Releases exist
    Releases have spaces between
    Release End - no space before
    Release Color exists
    Sprints exist
    Sprints have dates
    Sprints Color exist
    Sprint Marker 
    Notes exists
    Notes have dates
    Notes end exist */
    let rdata = this.getData();
    function whatsMissing() {
      this.releasesName = false,
        this.releasesNameLocation = "",
        this.releasesHaveBlanks = false,
        this.releasesEnd = false,
        this.releasesEndLocation = "",
        this.releasesHaveColor = false,
        this.sprintsName = false,
        this.sprintsNameLocation = "",
        this.sprintsHaveDates = false,
        this.sprintsEnd = false,
        this.sprintsEndLocation = "",
        this.sprintColors = false,
        this.notesName = false,
        this.notesNameLocation = "",
        this.notesEnd = false,
        this.notesEndLocation = "",
        this.notesDates = false
    }
    let check = new whatsMissing();
    let dataloc = rdata.findIndex(x => x[0] == DatabegingHere);
    let rcalendardata = rdata.slice(dataloc);
    rdata.forEach((x, y) => {
      if (y < dataloc + 1) {
        //skip these
      }
      else {
        switch (x[0]) {
          case SprintsName:
            check.sprintsName = true;
            check.sprintsNameLocation = y;
            if (x[3] !== "" && x[4] !== "") {
              check.sprintColors = true;
            }
            break;
          case SprintsEndof:
            check.sprintsEnd = true;
            check.sprintsEndLocation = y;
            break;
          case ReleasesName:
            check.releasesName = true;
            check.releasesNameLocation = y;
            break;
          case ReleasesEndof:
            check.releasesEnd = true;
            check.releasesEndLocation = y;
            break;
          case NotesName:
            check.notesName = true;
            check.notesNameLocation = y;
            break;
          case NotesEndof:
            check.notesEnd = true;
            check.notesEndLocation = y;
            break;
        }
      }
    })
    /* verify releases have blanks, and color */
    if (check.sprintsName && check.sprintsEnd) {
      // look for something in date field
      check.sprintsHaveDates = true;
      for (let i = check.sprintsNameLocation + 1; i < check.sprintsEndLocation; i++) {
        let x = rdata[i];
        if ((isNaN(Date.parse(x[1]))) || isNaN(Date.parse(x[2]))) {
          check.sprintsHaveDates = false;
        }
      }
    }
    if (check.notesName && check.notesEnd) {
      check.notesDates = true;
      for (let i = check.notesNameLocation + 1; i < check.notesEndLocation; i++) {
        let x = rdata[i];
        if (isNaN(Date.parse(x[1]))) {
          check.notesDates = false;
        }
      }
    }
    if (check.releasesName && check.releasesEnd) {
      check.releasesHaveBlanks = true;
      if (rdata[check.releasesEndLocation - 1][0] == "") {
        check.releasesHaveBlanks = false;
      }
      else {
        check.releasesHaveBlanks = false;
        // let oneblank = false;
        for (let i = check.releasesNameLocation + 1; i < check.releasesEndLocation; i++) {
          x = rdata[i];
          if (x[0] == "") {// check for at least one blank line between releases
            //   oneblank = true;
            check.releasesHaveBlanks = true;
          }
        }
      }
      //paul
      check.releasesHaveBlanks = true;
    }
    return check;
  }
  this.showValidationResults = function (check) {
    function showerror(missingItems) {
      //var markup = "There was a problem querying epics for team sheet "+teamsheet+ "\n\n   Query was \n\n" +query + "\n\n Return code " + error+ "\n\n Does this team actually have any epics?";
      let markup = "";
      missingItems.forEach(x => markup = markup + x + "\n")
      var ui = SpreadsheetApp.getUi();
      ui.alert("Oops", markup, ui.ButtonSet.OK);
    }
    /* check and see if the results are ok, if so just return */
    if (check.releasesName && check.releasesHaveBlanks && check.releasesEnd && check.sprintsName && check.sprintsHaveDates && check.sprintsEnd && check.notesName && check.notesEnd && check.notesDates) {
      return true;
    }
    else {
      let missingItems = [];
      if (!check.releasesName) {
        missingItems.push("Releases section could not be found - No 'Releases' in Col A in Configuration tab");
        check.releasesHaveBlanks = true;
      }
      if (!check.releasesEnd) {
        missingItems.push("End of Releases is missing - No 'End of Releases' in Col A in Configuration tab")
        check.releasesHaveBlanks = true;
      }
      if (!check.releasesHaveBlanks) {
        missingItems.push("No blank lines were found between releases in the Release section of the Configuration tab. Ensure that a blank line is present between Major release types")
      }
      if (!check.sprintsName) {
        missingItems.push("Sprints section could not be found - No 'Sprints' in Col A in Configuration tab");
        check.sprintsHaveDates = true;
      }
      if (!check.sprintsEnd) {
        missingItems.push("End of Sprints is missing - No 'End of Sprints' in Col A in Configuration tab");
        check.sprintsHaveDates = true;
      }
      if (!check.sprintsHaveDates) {
        missingItems.push("One or more sprint enteries is missing a date or contains an invalid date in Configuration tab")
      }
      if (!check.notesName) {
        missingItems.push("Notes section could not be found - No 'Notes' in Col A in Configuration tab")
        check.notesDates = true;
      }
      if (!check.notesEnd) {
        missingItems.push("End of Notes is missing - No 'End of Notes' in Col A in Configuration tab")
        check.notesDates = true;
      }
      if (!check.notesDates) {
        missingItems.push("One or more Notes enteries is missing a date or contains an invalid date in Configuration tab")
      }
      showerror(missingItems);
      return false;
    }
  }
}
var calendarsheet = function () {
  // const calendarpage = "Calendar";
  this.setCalendarsheet = function (sheetname) {
    SpreadsheetApp.setActiveSheet(sheet.getSheetByName(sheetname.title));
    actSheet = sheet.getActiveSheet();
  }
  this.setcalendarHeader = function (configstartrowcolumn, displayrow, rdata, writeout, rcolors) {
    function calendar_rec() {
      this.start,
        this.end,
        this.datewidth,
        this.spacerwidth,
        this.fontcolor,
        this.backgroundcolor
    }
    function calendar_header(calendarStartRow, calendarStartColumn, calendarStartDate, calendarEndDate) {
      function generateCalendar(calendarStartRow, calendarStartColumn, startDate, endDate) {
        const monthstitle = ['Jan', 'Feb', 'March', 'April', 'May', 'June', 'July', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec'];
        const weektitle = ["S", "M", "T", "W", "T", "F", "S"];
        let mbeginMonth = new Date(calendarStartDate).getMonth();
        let nextsevendays = new Date(calendarStartDate);
        nextsevendays.setDate(nextsevendays.getDate() + 6);
        let begindate = new Date(calendarStartDate);
        let calendar1titles = [];
        let calendar2titles = [];
        let calendar3titles = [];
        for (let i = 0; i < CalendarSheeetStartColumnOffet0; i++) {
          calendar1titles.push(""); // set a blank for the first column
          calendar2titles.push(""); //
          calendar3titles.push(""); // set a blank
        }
        let w = daysBetween(calendarStartDate, calendarEndDate) + 1; //was 2 
        let we = Math.round(w / 7);
        let i = 0;
        do {
          let mendMonth = new Date(nextsevendays).getMonth();
          let mbeginMonthName = monthstitle[mbeginMonth];
          let mendMonthName = monthstitle[mendMonth];
          let mtitle = mbeginMonthName;
          if (mbeginMonthName !== mendMonthName) {
            mtitle = mbeginMonthName + "/" + mendMonthName;
          }
          calendar1titles.push(mtitle);
          // Logger.log("Begin %s, end %s, Month title %s", begindate, sevendays, mtitle);
          begindate.setDate(begindate.getDate() + 7);
          nextsevendays.setDate(nextsevendays.getDate() + 7);
          mbeginMonth = new Date(begindate).getMonth();
          i++;
          for (let k = 0; k < 7; k++) { //push the padding needed for the sheet
            calendar1titles.push("");
          }
        } while (i < we);

        let s = new Date(calendarStartDate);
        let d = s.getDate();
        let dow = s.getDay();
        for (let ii = 1; ii <= w; ii++) {
          calendar2titles.push(weektitle[dow]);
          calendar3titles.push(d);
          s.setDate(s.getDate() + 1);
          d = s.getDate()
          dow = s.getDay();
          if (Number.isInteger(ii / 7)) {
            calendar2titles.push("")
            calendar3titles.push("")
          }
        }
        // synce the sizes of these with padding if necessary
        if (calendar1titles.length > calendar2titles.length) {
          let mc = calendar1titles.length - calendar2titles.length;
          for (let iii = 0; iii < mc; iii++) {
            calendar2titles.push("");
            calendar3titles.push("");
          }
        }
        if (calendar1titles.length < calendar2titles.length) {
          let mc = calendar2titles.length - calendar1titles.length;
          for (let iii = 0; iii < mc; iii++) {
            calendar1titles.push("");
          }
        }
        let titles = [];
        titles.push(calendar1titles);
        titles.push(calendar2titles);
        titles.push(calendar3titles);
        return titles;
      }
      const cstart = new Date(calendarStartDate).getTime();
      const cend = new Date(calendarEndDate).getTime();
      const calInfo = generateCalendar(calendarStartRow, calendarStartColumn, cstart, cend);
      return calInfo;
      //  Logger.log(weeks);
    }
    function writeCal(cal, displayrow, c) {
      function colors(fontcolor, backgroundcolor, rowlength, rows) {
        let fc = [];
        let bc = [];
        for (let i = 0; i < rows; i++) {
          let f = [];
          let b = [];
          for (let j = 0; j < rowlength; j++) {
            f.push(fontcolor);
            b.push(backgroundcolor);
          }
          fc.push(f);
          bc.push(b)
        }
        return { fc, bc };
      }
      let r = actSheet.getRange(displayrow, 1, cal.length, cal[1].length);
      let fbColors = colors(c.fontcolor, c.backgroundcolor, cal[0].length, cal.length);
      r.setValues(cal).setFontColors(fbColors.fc).setBackgrounds(fbColors.bc);
      //r.setFontColors(fbColors.fc);
      //r.setBackgrounds(fbColors.bc);
      //lets build an array fo "Center" values for aligments
      let carrayrow = new Array(cal[1].length).fill("Center");
      // carrayrow.forEach((x,y) => {carrayrow[y] = "Center" });
      let carray = new Array(cal.length).fill(carrayrow);
      //  carray.forEach((x,y) => {carray[y].push(carrayrow)});
      r.setHorizontalAlignments(carray);
      // set the month row styling
      r = actSheet.getRange(displayrow, 1, 1, cal[1].length);
      r.setFontWeight("Bold");
      r.setHorizontalAlignment("Center");
      //set constants to the left
      const weekheader = [["Months"], ["Days"], ["Date"]];
      let w = actSheet.getRange(displayrow, 1, 3);
      w.setValues(weekheader);
      w.setFontWeight("Normal");
      w.setHorizontalAlignment("Left");
      r.breakApart();
      //SpreadsheetApp.flush();
      let startc = CalendarSheetStartColumn;
      // merge ranges with month values
      for (let i = startc; i <= cal[1].length; i = i + 8) {
        let mth = actSheet.getRange(displayrow, i, 1, 7);
        // s.getRange(row+1,1,1,2).mergeAcross();
        mth.mergeAcross();
        // SpreadsheetApp.flush();
      }

      SpreadsheetApp.flush();
    }
    let c = new calendar_rec();
    let rc = decode(configstartrowcolumn)
    let m = rc.column
    let mm = rc.column - 1;
    let r = rc.row;
    c.start = rdata[rc.row][rc.column - 1];
    c.end = rdata[rc.row][rc.column];
    c.datewidth = rdata[rc.row][rc.column + 1];
    c.spacerwidth = rdata[rc.row][rc.column + 2];
    c.backgroundcolor = rcolors.backgrounddata[rc.row][rc.column - 1];
    c.fontcolor = rcolors.fontcolordata[rc.row][rc.column - 1];
    let calInfo = calendar_header(rc.row, rc.column, c.start, c.end);
    // Logger.log(titles)
    if (writeout) {
      writeCal(calInfo, displayrow, c);
    }
    return c;
  }
  this.setCalwidths = function (cal, displayrow) {
    //  SpreadsheetApp.setActiveSheet(sheet.getSheetByName(CalendarSheetName));
    //  actSheet = sheet.getActiveSheet();
    let r = actSheet.getDataRange();
    let lc = r.getLastColumn();
    for (let i = CalendarSheetStartColumn; i <= lc; i++) {
      let c = actSheet.getRange(displayrow + 2, i)
      let cv = c.getValue()
      if (cv !== "") {
        let cols = 7;
        if (i + cols > lc) {
          // Logger.log("stop")
          cols = lc - i + 1;
        }
        if (cols !== 0) {
          actSheet.setColumnWidths(i, cols, cal.datewidth);
        }
        i = i + 7;
        if (i < lc) {
          actSheet.setColumnWidth(i, cal.spacerwidth);
        }
      }
      else {
        actSheet.setColumnWidth(i, cal.spacerwidth);
      }
    }
  }
  this.setsprints = function (sprints, displayrow, cal) {
    let cdays = daysBetween(cal.start, cal.end) + 2;
    let cdayswPad = cdays + 1 + Math.round(cdays / 7); // this should be all cells on calendar row with pad between weeks
    // fill an empty calendar line
    let sprintline = [];
    for (let ii = 0; ii < CalendarSheetStartColumn; ii++) {
      sprintline.push("");
    }
    for (let ii = 0; ii < cdayswPad; ii++) {
      sprintline.push("");
    }
    // set sprint line to length of calendar
    sprints.forEach((c) => {
      let d = daysBetween(cal.start, c.begin);
      let sd = daysBetween(c.begin, c.end) + d + 1;
      //    Logger.log("Sprint %s, starts %s, and ends here %s", c.sprint, d, sd);
      let sp = c.sprint;
      for (let iii = d; iii < sd; iii++) {
        sprintline[iii] = sp;
        sp = CalendarSprintDayMarker;
      }
    })
    let lrow = actSheet.getLastColumn();
    // CalendarSheetStartRow was displayrow -1
    //let drow = actSheet.getRange(CalendarSheetStartRow, CalendarSheetStartColumn+1, 1, lrow)
    let drow = actSheet.getRange(CalendarSheetStartRow, 1, 1, lrow)
    let dvalues = drow.getValues();
    let sprintrow = [];
    let srow = [];
    let j = 0;
    dvalues[0].forEach((x, y) => {
      let m = x;
      if (y >= CalendarSheeetStartColumnOffet0) {
        if (x !== "") {
          sprintrow.push(sprintline[j]);
          j++;
        }
        else {
          sprintrow.push('');
        }
      }
      else {
        sprintrow.push('');
      }
    });
    srow.push(sprintrow);
    drow = actSheet.getRange(displayrow, 1, 1, sprintrow.length);
    drow.setValues(srow);
    let sprintbackgrounds = [];
    let sprintfontcolors = [];
    for (let i = 0; i < sprintrow.length; i++) {
      if ((sprintrow[i] == CalendarSprintDayMarker) || (sprintrow[i] == "")) {
        sprintbackgrounds.push(sprints[0].backgroundcolor);
        sprintfontcolors.push(sprints[0].fontcolor);
      }
      else {
        sprintbackgrounds.push(sprints[0].markerbackgroundcolor);
        sprintfontcolors.push(sprints[0].markerfontcolor);
      }
    }
    let sbc = [];
    sbc.push(sprintbackgrounds);
    drow.setBackgrounds(sbc);
    let sfc = [];
    sfc.push(sprintfontcolors);
    drow.setFontColors(sfc);
    let snamerow = actSheet.getRange(displayrow, 2, 1, 1);
    snamerow.setValue(SprintsName);
    srow[0].forEach((x, y) => {
      if ((x == CalendarSprintDayMarker) || (x == SprintsName) || (x == "")) {
        actSheet.getRange(displayrow, y + 1).setFontWeight("Normal").setHorizontalAlignment("center");
      }
      else {
        actSheet.getRange(displayrow, y + 1).setFontWeight("Bold").setHorizontalAlignment("center");
      }
    })
    SpreadsheetApp.flush();
    // Logger.log(sprintLocation, cal)

  }
  this.setreleases = function (releases, cal) {
    /* Read through the release data and build out rows for the spreadsheet
    create an additional array of font and background colors that can be used to color the releases
    */
    let cdays = daysBetween(cal.start, cal.end) + 2;
    //let cdayswPad = cdays + 1 + Math.round(cdays / 7); // this should be all cells on calendar row with pad between weeks
    // fill an empty calendar line
    function releaseDisplayRecord() {
      this.value = "",
        this.backgroundcolor = "#ffffff",
        this.fontcolor = "#000000"
    }
    let releasedisplayline = [];
    let lcolumn = actSheet.getLastColumn();
    for (let ii = 0; ii < lcolumn; ii++) {
      let rdl = new releaseDisplayRecord();
      releasedisplayline.push(rdl);
    }

    // set a placeholder array for release lines
    let releaseDisplayRecords = [];
    let lastreleaserow = releaseDisplayRecords.length
    let firstreleaserow = 0;
    let currentMajor = '';
    releases.forEach((x) => {
      // set the headers for release lines and build out blank lines for each env 
      if (currentMajor !== x.major) {
        currentMajor = x.major;
        let numoenvs = x.envs.length;
        let trdl = JSON.parse(JSON.stringify(releasedisplayline));
        releaseDisplayRecords.push([...trdl]);
        lastreleaserow = releaseDisplayRecords.length;
        firstreleaserow = lastreleaserow;
        for (let i = 0; i < numoenvs; i++) {
          // go through the environments for this release
          let trdl = JSON.parse(JSON.stringify(releasedisplayline));
          releaseDisplayRecords.push([...trdl]);
        }
        releaseDisplayRecords[lastreleaserow][0].value = currentMajor;
        releaseDisplayRecords[lastreleaserow][0].backgroundcolor = x.envs[0].majorcolor;
        releaseDisplayRecords[lastreleaserow][0].fontcolor = x.envs[0].majorfontcolor;

        for (let i = 0; i < x.envs.length; i++) {
          releaseDisplayRecords[lastreleaserow + i][1].value = x.envs[i].env;
          releaseDisplayRecords[lastreleaserow + i][1].backgroundcolor = x.envs[0].majorcolor;
          releaseDisplayRecords[lastreleaserow + i][1].fontcolor = x.envs[0].majorfontcolor;
          Logger.log("Major is %s i = %s, lastreleaserow = %s, value is %s, background is %s, fontcolor is %s", currentMajor, i, lastreleaserow, x.envs[i].env, x.backgroundcolor, x.fontcolor);
        }
        Logger.log("begins %s, ends %s", x.begin, x.end)
        lastreleaserow = releaseDisplayRecords.length;
      }
      let d = daysBetween(cal.start, x.begin);
      let sd = daysBetween(x.begin, x.end) + d + 2;
      // Logger.log("Major %s, version %s, env %s, begins at %s, ends at %s", x.major, x.version, x.env, x.begin, x.end);
      // Logger.log("version %s, starts %s, and is %s long", x.version, d+2, sd);
      // Logger.log("8 goes into %s, %s times",sd, sd%8)
      /**
       * Special case, if this is a 1 day event, so no begin marker
       */
      let eventlength = daysBetween(x.begin, x.end);
      // Logger.log("event length %s, began %s, ends %s", eventlength,x.begin,x.end)
      let sp;
      if (eventlength == 0) {
        sp = x.version;
      }
      else {
        sp = ReleaseBeginMarker + x.version;
      }
      let myrow = firstreleaserow + x.rowoffset;
      for (let iii = d + 2; iii <= sd; iii++) { // was d + 1
        releaseDisplayRecords[myrow][iii].value = sp
        releaseDisplayRecords[myrow][iii].backgroundcolor = x.backgroundcolor;
        releaseDisplayRecords[myrow][iii].fontcolor = x.fontcolor;
        sp = "";
        if (iii == sd) {
          //  releaserows[myrow][iii] = "|";
          /**
           * Special case, if releaseDisplayRecords[myrow][iii].value contains a value that means this is a 1 day event,
           * don't replace with release marker
           */
          if (!releaseDisplayRecords[myrow][iii].value) {
            releaseDisplayRecords[myrow][iii].value = ReleaseEndMarker;
          }
          releaseDisplayRecords[myrow][iii].backgroundcolor = x.backgroundcolor;
          releaseDisplayRecords[myrow][iii].fontcolor = x.fontcolor;
        }
      }
    });
    lcolumn = actSheet.getLastColumn();
    releaseDisplayRecords.forEach((x, y) => {
      let relcell = CalendarSheeetStartColumnOffet0; //was 0
      // insert a blank every 8 rows starting with the CalendarSheetStartColumn -1
      for (let i = CalendarSheeetStartColumnOffet0; i <= x.length; i++) { // was 1 and < not <=, then 0
        // Logger.log("x length %s", x.length);
        let jq = x;
        let any = i % 8; // was 
        if ((any == 0) && (i !== 0)) {
          let rdl = new releaseDisplayRecord();
          // Logger.log("relcell %s", relcell)
          //  Logger.log("Inserting a column at location %s, in row %s",relcell,y)
          /*
          rdl.backgroundcolor = releaseDisplayRecords[y][relcell - 1].backgroundcolor;
          rdl.fontcolor = releaseDisplayRecords[y][relcell - 1].fontcolor
          */
          rdl.backgroundcolor = releaseDisplayRecords[y][relcell].backgroundcolor;
          rdl.fontcolor = releaseDisplayRecords[y][relcell].fontcolor
          /****
           * Special Case, if adding a blank after the EndReleaseMarker then set the color to the next forward cell,
           * unless we are at the end of x (the end of the calendar). then forget it
           */
          if ((x[i].value == ReleaseEndMarker) && (i <= x.length)) {
            rdl.backgroundcolor = releaseDisplayRecords[y][relcell + 1].backgroundcolor;
            rdl.fontcolor = releaseDisplayRecords[y][relcell + 1].fontcolor
          }
          //  rdl.value = "x"
          releaseDisplayRecords[y].splice(relcell + 1, 0, rdl);
        }
        relcell++;
      }
      releaseDisplayRecords[y].splice(lcolumn, releaseDisplayRecords[y].length - lcolumn);
    });
    let releaseDisplayValues = [];
    let releasedisplayBackgrounds = [];
    let releasedisplayfonts = [];
    releaseDisplayRecords.forEach((x) => {
      let nl = [];
      let nb = [];
      let nf = [];
      x.forEach((d) => {
        nl.push(d.value);
        nb.push(d.backgroundcolor);
        nf.push(d.fontcolor);
      });
      releaseDisplayValues.push(nl);
      releasedisplayBackgrounds.push(nb)
      releasedisplayfonts.push(nf);
    })
    drow = actSheet.getRange(releases[0].displayrow, 1, releaseDisplayValues.length, lcolumn);
    drow.setValues(releaseDisplayValues);
    drow.setBackgrounds(releasedisplayBackgrounds);
    drow.setFontColors(releasedisplayfonts);

    SpreadsheetApp.flush();
  }
  this.setMajorsVertical = function (displayrow) {
    // Logger.log("What do we have");
    let mlastRow = actSheet.getLastRow();
    let mrange = actSheet.getRange(displayrow + 1, 1, mlastRow, 2);
    let mdata = mrange.getValues();
    //mrange.mergeVertically();
    function majorlocation() {
      this.majorname,
        this.majoroffset,
        this.majorrows
    }
    let major = mdata[0][0];
    let mrows = 1;
    let majorArray = [];
    let first = true;
    let m = new majorlocation();
    m.majorname = mdata[0][0];
    m.majoroffset = 0;
    m.majorrows = 1;
    // mdata.shift();
    // get new majors and offset
    mdata.forEach((x, y) => {
      if (x[0] !== major) {
        if (x[0] !== "") {
          m.majorrows = mrows;
          majorArray.push(m);
          mrows = 1;
          major = x[0];
          m = new majorlocation();
          m.majorname = major;
          m.majoroffset = y;
          m.majorrows = mrows;
        }
        else {
          if (x[1] !== "") {
            mrows++;
          }
        }
      }
    });
    m.majorrows = mrows;
    majorArray.push(m);
    // lets go vertical on these locations
    let startrow = displayrow + 1;
    /* at startrow, use the majoroffset to get the top row, use majorrows to get the range of rows below
  
    */
    mrange.breakApart();
    majorArray.forEach(x => {
      let p = startrow + x.majoroffset;
      let mj = actSheet.getRange(startrow + x.majoroffset, 1, x.majorrows, 1)
      if (x.majorrows > 1) {
        mj.mergeVertically().setTextRotation(90).setFontWeight("bold").setHorizontalAlignment('center');
      }
      else {
        mj.setTextRotation(0).setFontWeight("bold").setHorizontalAlignment('center');
      }
    })
  }

  this.setnote = function (notes, displayrow, cal, sheetname) {
    SpreadsheetApp.setActiveSheet(sheet.getSheetByName(sheetname.title));
    actSheet = sheet.getActiveSheet();
    let rdata = actSheet.getDataRange().getValues();
    actSheet.clearNotes();
    notes.forEach((x, y) => {
      let cdays = daysBetween(cal.start, x.NoteDate);
      let cd = Math.floor(cdays / 7);
      let cdayswPad = cdays + 2 + cd; // this should be all cells on calendar row with pad between weeks
      const options = {
        year: "numeric",
        month: "2-digit",
        day: "numeric"
      };
      let nd = new Date(x.NoteDate).toLocaleString("en", options);
      let nt = nd + " - " + x.NoteText;
      actSheet.getRange(displayrow, cdayswPad + 1).setNote(nt).setBackground(x.NoteBackgroundColor).setFontColor(x.NoteFontColor);
      //  Logger.log("At location %s, this many blanks %s, this date %s",cdayswPad, cd,x.NoteDate);
      // Logger.log("calendar date is %s,this note %s",rdata[2][cdayswPad],x.NoteText);
    })
  }
  this.setfrozen = function (column, row) {
    actSheet.setFrozenColumns(0);
    actSheet.setFrozenRows(0);
    actSheet.setFrozenColumns(column);
    actSheet.setFrozenRows(row);
  }
  this.checksheetsexistance = function (sheetname) {
    let itt = sheet.getSheetByName(sheetname);
    let status = "";
    if (!itt) {
      sheet.insertSheet(sheetname);
      status = false;
    }
    else {
      status = true;
    }
    return status;
  }
}
function resetcolumns() {
  const rdata = new readconfig().getData();
  const rcolors = new readconfig().getBackgroundsandFontColors();
  const artifacts = new readconfig().getArtifacts(rdata);
  const calLocation = artifacts[artifacts.findIndex((e) => e.artifact == CalendarArtifcatName)];
  const calsheetname = artifacts[artifacts.findIndex((e) => e.artifact == CalendarTabArtifactName)];
  const calsheet = new calendarsheet();
  calsheet.setCalendarsheet(calsheetname);
  const cal = calsheet.setcalendarHeader(calLocation.configurationrowcolumn, calLocation.displayrow, rdata, false, rcolors);
  calsheet.setCalwidths(cal, calLocation.displayrow);
}
function main() {
  let rconfig = new readconfig();
  /**
   * Lets read the config and see if things look like they are set correctly
   */
  let check = rconfig.verifyConfig();
  const oktoproceed = rconfig.showValidationResults(check);
  if (!oktoproceed) {
    return 0;
  }
  /** 
   * Let's read in the configuration tab, then get the locations of some items there and
   * get background and font colors
   */
  const rdata = rconfig.getData();
  const locationsSRN = rconfig.artifactlocationsSRN();
  const rcolors = rconfig.getBackgroundsandFontColors();
  /**
   * Let's find the locations of some information used to build the calendar - from the Configuration tab
   */
  const artifacts = rconfig.getArtifacts(rdata);
  const calLocation = artifacts[artifacts.findIndex((e) => e.artifact == CalendarArtifcatName)];
  const calsheetname = artifacts[artifacts.findIndex((e) => e.artifact == CalendarTabArtifactName)];
  const notesLocation = artifacts[artifacts.findIndex((e) => e.artifact == NotesName)];
  /**
   * Let's build the calendar portion of the Release sheet
   */
  const calsheet = new calendarsheet();
  const sheetstatus = calsheet.checksheetsexistance(calsheetname.title);
  calsheet.setCalendarsheet(calsheetname);
  const cal = calsheet.setcalendarHeader(calLocation.configurationrowcolumn, calLocation.displayrow, rdata, true, rcolors);
  /**
   * Now let's get the release(s) information, Sprint information, and notes from the configuration tab
   */
  const releases = rconfig.getReleases(locationsSRN, artifacts, rdata, rcolors);
  const sprints = rconfig.getSprints(locationsSRN, rdata, rcolors);
  const notes = rconfig.getNotes(locationsSRN, rdata, rcolors);
  /**
   * now let's write out the Calendar, notes, and sprint information on the calendar
   */
  calsheet.setCalendarsheet(calsheetname);
  calsheet.setnote(notes, notesLocation.displayrow, cal, calsheetname);
  const sprintLocation = artifacts[artifacts.findIndex((e) => e.artifact == SprintsName)];
  calsheet.setsprints(sprints, sprintLocation.displayrow, cal);
  /** 
   * now let's build releases and write them to the calendar
   */
  const releaseLocation = artifacts[artifacts.findIndex((e) => e.artifact == ReleasesName)];
  calsheet.setreleases(releases, cal);
  /**
   * Let's set some text vertical for Release types
   */
  calsheet.setMajorsVertical(releaseLocation.displayrow);
  calsheet.setfrozen(CalendarSheetStartColumn - 1, sprintLocation.displayrow);
  /**
   * If this is the first time we are creating this sheet, let's set the widths of the rows
   */
  if (!sheetstatus) {
    calsheet.setCalwidths(cal, calLocation.displayrow);
  }
}

