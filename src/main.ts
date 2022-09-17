const configSheetName = "設定";
const AnniversarySheetName = "記念日";
const calendarIdCell = "B2";
const discordWebHookUrl = "B1";

const getSheetByName = (name: string) =>
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);

const getConfigSheet = () => getSheetByName(configSheetName);
const getAnniversarySheet = () => getSheetByName(AnniversarySheetName);

const getConfigCell = (cell: string) =>
  getConfigSheet()?.getRange(cell).getValue();

const requestMessage = (usename: string, content: string) => {
  UrlFetchApp.fetch(getConfigCell(discordWebHookUrl), {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ usename, content }),
  });
};

const notifyAnniversary = () => {
  const sheet = getAnniversarySheet();
  if (!sheet) return;
  const lastRowNumber: number = sheet.getLastRow();
  const today = getToday();
  const nextDay = getNextDay();
  for (let i = 2; i <= lastRowNumber; i++) {
    const cells = sheet.getRange(i, 1, 1, 4).getValues()[0];
    if (!cells) return;
    const [type, date, desc]: any = cells.slice(0, 3);
    const anniversary = new Date(date);
    const differenceDate = getDateDifference(anniversary, today);
    if (type === "記念日" && differenceDate % 100 === 0) {
      requestMessage(
        "記念日",
        `今日は ${desc} から ${differenceDate} 日目です!`
      );
    }

    if (
      today.getDate() === anniversary.getDate() &&
      today.getMonth() === anniversary.getMonth()
    ) {
      if (type === "記念日") {
        requestMessage(
          "記念日",
          `今日は ${desc} ${
            today.getFullYear() - anniversary.getFullYear()
          } 年目です!`
        );
      }
      if (type === "誕生日") {
        requestMessage(
          "誕生日",
          `今日は ${desc} の ${
            today.getFullYear() - anniversary.getFullYear()
          } 才の誕生日です!`
        );
      }
    }

    if (
      nextDay.getDate() === anniversary.getDate() &&
      nextDay.getMonth() === anniversary.getMonth()
    ) {
      if (type === "記念日") {
        requestMessage(
          "記念日",
          `明日は ${desc} ${
            nextDay.getFullYear() - anniversary.getFullYear()
          } 年目です!`
        );
      }
      if (type === "誕生日") {
        requestMessage(
          "誕生日",
          `明日は ${desc} の ${
            nextDay.getFullYear() - anniversary.getFullYear()
          } 才の誕生日です!`
        );
      }
    }
  }
};

const getToday = (): Date => new Date(new Date().toDateString());

const getNextDay = (): Date => {
  const nextDay = getToday();
  nextDay.setDate(nextDay.getDate() + 1);
  return nextDay;
};

const getDateDifference = (from: Date, to: Date): number => {
  const ms = to.getTime() - from.getTime();
  return Math.floor(ms / (1000 * 60 * 60 * 24));
};

const notifyWeekEvent = () => {
  const today = new Date();
  const calendar = CalendarApp.getCalendarById(getConfigCell(calendarIdCell));
  const nextDayStart = new Date(today.toDateString());
  nextDayStart.setDate(nextDayStart.getDate() + 1);

  const nextDayEnd = new Date(nextDayStart.toDateString());
  nextDayEnd.setDate(nextDayEnd.getDate() + 7);

  const nextDayEnd2 = new Date(nextDayEnd.toDateString());
  nextDayEnd2.setDate(nextDayEnd2.getDate() - 1);

  const events = calendar.getEvents(nextDayStart, nextDayEnd);

  const msg =
    events.length === 0
      ? "ありません"
      : events
          .map((e) => {
            const title = e.getTitle();
            const date = parseToMonthDate(e.getStartTime());
            return `  ${date}: ${title} `;
          })
          .join("\n");
  requestMessage(
    "予定",
    `来週 (${parseToMonthDate(nextDayStart)} - ${parseToMonthDate(
      nextDayEnd2
    )}) の予定:\n${msg}`
  );
};

const notifyNextDateEvent = () => {
  const today = new Date();
  const calendar = CalendarApp.getCalendarById(getConfigCell(calendarIdCell));
  const nextDayStart = new Date(today.toDateString());
  nextDayStart.setDate(nextDayStart.getDate() + 1);

  const nextDayEnd = new Date(nextDayStart.toDateString());
  nextDayEnd.setDate(nextDayStart.getDate() + 1);

  const events = calendar.getEvents(nextDayStart, nextDayEnd);

  const msg =
    events.length === 0
      ? "ありません"
      : events
          .map((e) => {
            const title = e.getTitle();
            const startTime = parseToTime(e.getStartTime());
            const endTime = parseToTime(e.getEndTime());
            return `${title}: ${startTime} - ${endTime}`;
          })
          .join("\n");
  requestMessage(
    "予定",
    `明日 (${parseToMonthDate(nextDayStart)}) の予定:\n${msg}`
  );
};

const parseToMonthDate = (date: GoogleAppsScript.Base.Date): string =>
  `${("00" + (date.getMonth() + 1)).slice(-2)}/${("00" + date.getDate()).slice(
    -2
  )}`;

const parseToTime = (date: GoogleAppsScript.Base.Date): string =>
  `${("00" + date.getHours()).slice(-2)}:${("00" + date.getMinutes()).slice(
    -2
  )}`;
