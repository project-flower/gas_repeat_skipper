/** カレンダー ID */
const calendarId = ""; // *****@gmail.com 等 (省略した場合、アカウント既定のカレンダー)
/** 開始日 */
const startDate = "2026/1/1";
/** 終了日 */
const endDate = "2026/12/31";

/**
 * 毎週土日を除く、祝日の場合はキャンセルするスケジュールを作成します。
 */
function schedule1() {
  main(
    calendarId,
    startDate,
    endDate,
    { span: "weekly" },
    { weekend: true, holiday: true },
    "cancel",
    "Test Schedule 1",
    "毎週土日を除く、祝日の場合はキャンセルします。",
  );
}

/**
 * 毎週 月・水・金曜日又は祝日の場合はキャンセルするスケジュールを作成します。
 */
function schedule2() {
  main(
    calendarId,
    startDate,
    endDate,
    { span: "weekly", weekdays: [1, 3, 5] },
    { holiday: true },
    "cancel",
    "Test Schedule 2",
    "毎週 月・水・金曜日又は祝日の場合はキャンセルします。",
  );
}

/**
 * 毎月27日又は土日・祝日の場合は次の平日に設定するスケジュールを作成します。
 */
function schedule3() {
  main(
    calendarId,
    "2026/1/27",
    endDate,
    { span: "monthly" },
    { weekend: true, holiday: true },
    "next",
    "Test Schedule 3",
    "毎月27日又は土日・祝日の場合は次の平日に設定します。",
  );
}

/**
 * 毎年12月25日又は土日・祝日の場合は前の平日に設定するスケジュールを作成します。
 */
function schedule4() {
  main(
    calendarId,
    "2026/12/25",
    "2031/12/31",
    { span: "yearly" },
    { weekend: true, holiday: true },
    "prev",
    "Test Schedule 4",
    "毎年12月25日又は土日・祝日の場合は前の平日に設定します。",
  );
}
