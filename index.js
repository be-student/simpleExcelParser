var XLSX = require("xlsx");
function finderA(sheet, index) {
  const temp = sheet["A" + index];
  return temp?.w;
}
function finderB(sheet, index) {
  const temp = sheet["B" + index];
  return temp?.v;
}
function parseNumber(sheet, alphabet, index) {
  return sheet[alphabet + index]?.v;
}
function parseDate(dateString, time) {
  if (typeof time != "string") {
    return;
  }
  let timeArray;
  dateArray = dateString.split("-");
  timeArray = time.split(":");
  return new Date(dateArray[0], dateArray[1], dateArray[2], timeArray[0]);
}
function parseDiff(timeString) {
  if (typeof timeString != "string") {
    return;
  }
  let tokenizer;
  if (timeString.includes("-")) {
    tokenizer = "-";
  }
  if (timeString.includes("~")) {
    tokenizer = "~";
  }
  if (tokenizer == undefined) {
    return;
  }

  let timeArray = timeString.split(tokenizer);
  if (!timeArray[0].includes(":") || !timeArray[1].includes(":")) {
    return;
  }
  let time1 = timeArray[0].split(":");
  let time2 = timeArray[1].split(":");

  let start = +time1[0] * 60 + +time1[1];
  let end = +time2[0] * 60 + +time2[[1]];
  return end - start;
}
async function parser() {
  const workbook = XLSX.readFile(process.argv[2]);
  const sheet = workbook.Sheets["코칭일지"];
  let maxDate = parseDate(finderA(sheet, 2), finderB(sheet, 2));
  let startIndex = 1;
  parseNumber(sheet, "E", "30");

  // console.log(sheet["E18"]);
  let feeTime = 0;
  let freeTime = 0;
  let coachTime = 0;

  while (true) {
    startIndex++;
    const A = finderA(sheet, startIndex);
    const B = finderB(sheet, startIndex);

    if (!A && finderA(sheet, startIndex + 1) == undefined) {
      break;
    }
    if (!A) {
      console.log("A" + startIndex + "가 비어있습니다");
      continue;
    }

    if (!B) {
      console.log("B" + startIndex + "가 비어있습니다");
      continue;
    }
    const newDate = parseDate(A, B);
    if (!newDate) {
      console.log(startIndex + "번째 줄의 " + "시간을 확인해주세요");
    }
    if (maxDate > parseDate(A, B)) {
      console.log("A" + startIndex + "시간을 확인해 주세요");
      continue;
    }
    maxDate = newDate;
    const feeMin = parseNumber(sheet, "D", startIndex);
    const freeMin = parseNumber(sheet, "E", startIndex);
    const coachMin = parseNumber(sheet, "F", startIndex);

    const diff = parseDiff(B);
    if (!diff) {
      console.log(startIndex + "줄의 시간을 계산할 수 없습니다");
      continue;
    }
    if (feeMin) {
      if (feeMin > 120) {
        console.log(startIndex + "번째 줄의 유료 시간이 120분을 초과했습니다");
      }
      if (diff != feeMin) {
        console.log(startIndex + "번째 줄의 유료 시간이 올바르지 않습니다");
      }
      feeTime += feeMin;
      continue;
    }
    if (freeMin) {
      if (freeMin > 120) {
        console.log(startIndex + "번째 줄의 무료 시간이 120분을 초과했습니다");
      }
      if (diff != freeMin) {
        console.log(startIndex + "번째 줄의 무료 시간이 올바르지 않습니다");
      }
      freeTime += freeMin;
      continue;
    }
    if (coachMin) {
      if (coachMin > 120) {
        console.log(
          startIndex + "번째 줄의 코더코 시간이 120분을 초과했습니다"
        );
      }
      if (diff * 2 > 120) {
        if (120 != coachMin) {
          console.log(startIndex + "번째 줄의 코더코 시간이 올바르지 않습니다");
        }
      } else {
        if (diff * 2 != coachMin) {
          console.log(startIndex + "번째 줄의 코더코 시간이 올바르지 않습니다");
        }
      }
      coachTime += coachMin;
      continue;
    }
  }

  console.log(feeTime + "분의 유료 시간이 있습니다");
  console.log(freeTime + "분의 무료 시간이 있습니다");
  console.log(coachTime + "분의 코더코 시간이 있습니다");
}
parser();
