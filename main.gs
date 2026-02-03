/** 日本の祝日 カレンダー */
const JapaneseHolidayCalendar = CalendarApp.getCalendarById(
  "ja.japanese#holiday@group.v.calendar.google.com",
);

/**
 * 移動先の日付を探します。
 * @param date スキップする日付
 * @param moveTo スキップする方向
 * @param skips スキップするオプション
 * @param minDate 設定できる最も前の日付
 * @param maxDate 設定できる最も先の日付
 * @returns 移動先となる日付を返します。
 */
function findMovedDay(date, moveTo, skips, minDate, maxDate) {
  let result = new Date(date);
  const minTime = minDate.getTime();
  const maxTime = maxDate.getTime();

  do {
    result.setDate(result.getDate() + (moveTo === "prev" ? -1 : 1));
    const time = result.getTime();

    if (time < minTime || maxTime < time) {
      return undefined;
    }

    const day = result.getDay();

    if (
      (skips?.holiday && isHoliday(result)) ||
      (skips?.weekend && isWeekend(day)) ||
      skips?.weekdays?.includes(day)
    ) {
      continue;
    }

    return result;
  } while (true);
}

/**
 * 日本の祝日かどうかを判定します。
 * @param date 判定する日付
 * @returns 祝日の場合、true を返します。
 */
function isHoliday(date) {
  // この方法だと七夕やクリスマスも true を返してしまう
  // return JapaneseHolidayCalendar.getEventsForDay(date).length > 0;
  const result = JapaneseHolidayCalendar.getEventsForDay(date);
  return result.length > 0 && result[0].getDescription() === "祝日";
}

/**
 * 土曜日もしくは日曜日かどうかを判定します。
 * @param day 判定する曜日
 * @returns 土曜日もしくは日曜日の場合、true を返します。
 */
function isWeekend(day) {
  return [0, 6].includes(day);
}

/**
 * メイン
 * @param calendarId カレンダー ID (Google メールアドレス 等、省略した場合、アカウントの既定のカレンダー)
 * @param start 開始日
 * @param end 終了日
 * @param cycle 周期
 * @param skips スキップするオプション
 * @param moveTo スキップ先
 * @param title スケジュール名
 * @param description スケジュールの説明
 */
function main(
  calendarId,
  start,
  end,
  cycle,
  skips,
  moveTo,
  title,
  description,
) {
  const calendarId_ = calendarId || "primary";
  const calendar = CalendarApp.getCalendarById(calendarId_);

  if (!calendar) {
    throw new Error(`カレンダー ${calendarId_} が見つかりません。`);
  }

  let startDate, lastDate;

  try {
    startDate = new Date(start);
  } catch (e) {
    throw new Error(`start "${start}" は日付に変換できません。`);
  }

  try {
    lastDate = new Date(end);
  } catch (e) {
    throw new Error(`end "${end}" は日付に変換できません。`);
  }

  if (lastDate < startDate) {
    throw new Error(`${start} ～ ${end} の方向は矛盾します。`);
  }

  if (!cycle) {
    throw new Error(`cycle を指定してください。`);
  }

  const span = cycle.span;

  switch (span) {
    case "weekly":
    case "monthly":
    case "yearly":
      break;
    default:
      throw new Error(
        "cycle.span は weekly monthly yearly のいずれかを指定してください。",
      );
  }

  if (skips?.holiday && skips.holiday !== true) {
    throw new Error("skips.holiday は true を指定するか、省略してください。");
  }

  if (cycle.weekdays && cycle.weekdays.length !== undefined) {
    if (span !== "weekly") {
      throw new Error(
        "cycle.weekdays は cycle.span が weekly 以外では指定できません。",
      );
    }

    for (const weekday of cycle.weekdays) {
      if (weekday < 0 || 6 < weekday) {
        throw new Error("cycle.weekdays は 0-6 の配列を指定してください。");
      }

      if (
        skips?.weekdays &&
        skips.weekdays.includes !== undefined &&
        skips.weekdays.includes(weekday)
      ) {
        throw new Error(
          `cycle.weekdays と skips.weekdays の要素 ${weekday} が重複しています。`,
        );
      }
    }
  }

  if (skips?.weekdays && skips.weekdays.length !== undefined) {
    for (const weekday of skips.weekdays) {
      if (weekday < 0 || 6 < weekday) {
        throw new Error("skips.weekdays は 0-6 の配列を指定してください。");
      }
    }
  }

  if (skips?.weekend && skips.weekend !== true) {
    throw new Error("skips.weekend は true を指定するか、省略してください。");
  }

  if (!["cancel", "next", "prev"].includes(moveTo)) {
    throw new Error(
      "moveTo は prev 又は next 又は cancel を指定してください。",
    );
  }

  if (!title) {
    throw new Error("title を指定してください。");
  }

  const recurrence = CalendarApp.newRecurrence();
  let count = 0;
  let actualStartDate;

  for (let date = new Date(startDate); date <= lastDate; ) {
    const day = date.getDay();
    let skip = false;
    let exclude = false;

    if (cycle.weekdays && !cycle.weekdays.includes(day)) {
      skip = true;
    } else {
      if (
        (skips?.holiday && isHoliday(date)) ||
        skips?.weekdays?.includes(day)
      ) {
        exclude = true;
      }

      if (skips?.weekend && isWeekend(day)) {
        exclude = true;
      }
    }

    if (exclude && moveTo === "cancel") {
      skip = true;
    }

    if (!skip) {
      const resultDate = exclude
        ? findMovedDay(date, moveTo, skips, startDate, lastDate)
        : date;

      if (resultDate) {
        if (!actualStartDate) {
          actualStartDate = resultDate;
        }

        ++count;
        recurrence.addDate(resultDate);
      }
    }

    switch (cycle.span) {
      case "weekly":
        date.setDate(date.getDate() + 1);
        break;
      case "monthly":
        date.setMonth(date.getMonth() + 1);

        {
          // 次の月
          const month = new Date(date).getMonth();
          date.setDate(startDate.getDate());

          if (date.getMonth() !== month) {
            // 次の月に同じ日付がなければ、月末を設定する
            date.setMonth(month + 1, 0);
          }
        }

        break;
      case "yearly": {
        const month = startDate.getMonth();
        const next = new Date(
          date.getFullYear() + 1,
          month,
          startDate.getDate(),
        );

        if (next.getMonth() !== startDate.getMonth()) {
          next.setFullYear(date.getFullYear() + 1, month + 1, 0);
        }

        date = new Date(next);
      }
    }
  }

  if (!actualStartDate) {
    throw new Error("作成できる日付がありません。");
  }

  if (count === 1) {
    // 1 日のみの EventRecurrence は 2 日作成されてしまうため、実行しない
    throw new Error("作成できる日付が 1 日しかありません。");
  }

  const events = calendar.createAllDayEventSeries(
    title,
    actualStartDate,
    recurrence,
    {
      description,
    },
  );
}
