/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("run").onclick = runWord;
  } else if (info.host === Office.HostType.Excel) {
    document.getElementById("run").onclick = runExcel;
  }
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("run").onclick = runPowerPoint;
  }
});

export async function runExcel() {
  return Excel.run(async (context) => {
    var originalRange = context.workbook.getSelectedRange();
    originalRange.load();
    await context.sync();

    var isSet = false;

    if (originalRange.text.length != 0) {
      fetch(window.location.origin + "/assets/abbs.json")
        .then((res) => res.json())
        .then((json) => {
          json.forEach(function (data) {
            console.log(originalRange.text[0][0]);
            if (data.abb === originalRange.text[0][0].trim().toUpperCase()) {
              isSet = true;
              document.getElementById("abbrevationID").innerText = "Explanation: " + data.explanation;
              return;
            }
          });
        });
    } //

    if (isSet === false) document.getElementById("abbrevationID").innerText = "Explanation not found";
    await context.sync();
  });
}

export async function runPowerPoint() {
  return PowerPoint.run(async (context) => {
    var originalRange = context.presentation.getSelectedTextRange();
    originalRange.load("text");
    await context.sync();
    var isSet = false;

    if (originalRange.text.length != 0) {
      fetch(window.location.origin + "/assets/abbs.json")
        .then((res) => res.json())
        .then((json) => {
          json.forEach(function (data) {
            if (data.abb === originalRange.text.trim().toUpperCase()) {
              isSet = true;
              document.getElementById("abbrevationID").innerText = "Explanation: " + data.explanation;
              return;
            }
          });
        });
    } //

    if (isSet === false) document.getElementById("abbrevationID").innerText = "Explanation not found";
    await context.sync();
  });
}

export async function runWord() {
  return Word.run(async (context) => {
    var originalRange = context.document.getSelection();
    originalRange.load("text");
    await context.sync();

    var isSet = false;

    if (originalRange.text.length != 0) {
      fetch(window.location.origin + "/assets/abbs.json")
        .then((res) => res.json())
        .then((json) => {
          json.forEach(function (data) {
            if (data.abb === originalRange.text.trim().toUpperCase()) {
              isSet = true;
              document.getElementById("abbrevationID").innerText = "Explanation: " + data.explanation;
              return;
            }
          });
        });
    } //

    if (isSet === false) document.getElementById("abbrevationID").innerText = "Explanation not found";
    await context.sync();
  });
}
