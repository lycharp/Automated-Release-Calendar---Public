function myFunction() {
  let x = new Date(2022,0,1);
  let d1 = new Date(2022, 0, 1);
  let d2 = new Date(2022, 12, 0); 
  let m = daysBetween(d1,d2);
  Logger.log(m)
}
 function daysBetween(date1, date2) {

 // adjust diff for for daylight savings
 var hoursToAdjust = Math.abs(date1.getTimezoneOffset() /60) - Math.abs(date2.getTimezoneOffset() /60);
 // apply the tz offset
 date2.addHours(hoursToAdjust); 

    // The number of milliseconds in one day
    var ONE_DAY = 1000 * 60 * 60 * 24

    // Convert both dates to milliseconds
    var date1_ms = date1.getTime()
    var date2_ms = date2.getTime()

    // Calculate the difference in milliseconds
    var difference_ms = Math.abs(date1_ms - date2_ms)

    // Convert back to days and return
    return Math.round(difference_ms/ONE_DAY)

}

// you'll want this addHours function too 

Date.prototype.addHours= function(h){
    this.setHours(this.getHours()+h);
    return this;
}
function decode(cell) {
  const fromA1Notation = (cell) => {
    const [, columnName, rown] = cell.toUpperCase().match(/([A-Z]+)([0-9]+)/);
    //   const [, columnName, row] = cell.toUpperCase().match(/([A-Z]+)([0-9]+)/);
    const characters = 'Z'.charCodeAt() - 'A'.charCodeAt() + 1;
    let row = Number(rown);
    let column = 0;
    columnName.split('').forEach((char) => {
      column *= characters;
      column += char.charCodeAt() - 'A'.charCodeAt() + 1;
    });

    return { row, column };
  };
  let rr = fromA1Notation(cell);
  return rr;
}