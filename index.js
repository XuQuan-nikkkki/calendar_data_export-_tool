const { Solar, Lunar, HolidayUtil } = require("lunar-javascript");

const readline = require("readline").createInterface({
  input: process.stdin,
  output: process.stdout,
});

const getNextMonthRange = () => {
  const now = new Date();

  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth();

  const nextYear = currentYear + (currentMonth === 11 ? 1 : 0);
  const nextMonth = currentMonth === 11 ? 0 : currentMonth + 1;

  // 设置时区偏移
  const firstDay = new Date(Date.UTC(nextYear, nextMonth, 1));
  const lastDay = new Date(Date.UTC(nextYear, nextMonth + 1, 0));

  return {
    start: firstDay.toISOString().split("T")[0],
    end: lastDay.toISOString().split("T")[0],
  };
};

const isValidDate = (dateStr) => {
  if (!dateStr) return true;
  const regex = /^\d{4}(-|\/)(0[1-9]|1[0-2])(-|\/)?(0[1-9]|[12]\d|3[01])?$/;
  if (!regex.test(dateStr)) return false;

  const parts = dateStr.split(/[-\/]/);
  if (parts.length < 2) return false;

  const year = parseInt(parts[0]);
  const month = parseInt(parts[1]);
  const day = parts[2] ? parseInt(parts[2]) : 1;

  const date = new Date(year, month - 1, day);
  return (
    date.getFullYear() === year &&
    date.getMonth() === month - 1 &&
    date.getDate() === day
  );
};

const formatDate = (dateStr) => {
  if (!dateStr) return "";
  const parts = dateStr.split(/[-\/]/);
  if (parts.length === 2) {
    return `${parts[0]}-${parts[1]}-01`;
  }
  return dateStr.replace(/\//g, "-");
};

const nextMonth = getNextMonthRange();
let startDate = "";
let endDate = "";

const askStartDate = () => {
  readline.question(`请输入开始日期 (默认 ${nextMonth.start}): `, (input) => {
    if (!input) {
      startDate = nextMonth.start;
      askEndDate();
    } else if (isValidDate(input)) {
      startDate = formatDate(input);
      askEndDate();
    } else {
      console.log(
        "日期格式不正确，请使用 YYYY-MM-DD 或 YYYY-MM 或 YYYY/MM/DD 或 YYYY/MM 格式"
      );
      askStartDate();
    }
  });
};

const askEndDate = () => {
  readline.question(`请输入结束日期 (默认 ${nextMonth.end}): `, (input) => {
    if (!input) {
      endDate = nextMonth.end;
      confirmAndProceed();
    } else if (isValidDate(input)) {
      endDate = formatDate(input);
      confirmAndProceed();
    } else {
      console.log(
        "日期格式不正确，请使用 YYYY-MM-DD 或 YYYY-MM 或 YYYY/MM/DD 或 YYYY/MM 格式"
      );
      askEndDate();
    }
  });
};

const data = [];
const generateDataForDate = (date) => {
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const day = date.getDate();

  const solarDate = Solar.fromYmd(year, month, day);
  const solarFestivals = solarDate.getFestivals();
  const solarOtherFestivals = solarDate.getOtherFestivals();

  const lunarDate = Lunar.fromDate(new Date(year, month - 1, day));
  const lunarFestivals = lunarDate.getFestivals();
  const lunarOtherFestivals = lunarDate.getOtherFestivals();
  const jieqi = lunarDate.getJieQi();
  const yi = lunarDate.getDayYi();
  const ji = lunarDate.getDayJi();

  const holiday = HolidayUtil.getHoliday(year, month, day);
  const isWork = holiday?.isWork() ? "调休" : "";

  const dataItem = {
    日期: date.toISOString().split("T")[0],
    事件: "高能量事件",
    节日: [
      jieqi,
      ...solarFestivals,
      ...solarOtherFestivals,
      ...lunarFestivals,
      ...lunarOtherFestivals,
    ]
      .join(" ")
      .trim(),
    调休: isWork,
    宜: yi.join(", "),
    忌: ji.join(", "),
  };
  data.push(dataItem);
};

const generateData = (start, end) => {
  const firstDate = new Date(start);
  const lastDate = new Date(end);

  let currentDate = firstDate;
  while (currentDate <= lastDate) {
    generateDataForDate(currentDate);
    currentDate.setDate(currentDate.getDate() + 1);
  }

  // 输出生成的事件数据
  const XLSX = require("xlsx");
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(workbook, worksheet, "事件数据");

  const defaultFileName = `Calendar_Data_${startDate}_${endDate}.xlsx`;
  const downloadsPath = require("path").join(
    require("os").homedir(),
    "Downloads",
    defaultFileName
  );
  XLSX.writeFile(workbook, downloadsPath);
};

const confirmAndProceed = () => {
  readline.question(
    `确认要生成从 ${startDate} 到 ${endDate} 之间的数据吗? (Y/N) `,
    (answer) => {
      if (answer.toLowerCase() === "y") {
        console.log("继续执行脚本...");
        generateData(startDate, endDate);
        console.log(`Excel 文件已生成至: ${require("path").join(require("os").homedir(), "Downloads", `Calendar_Data_${startDate}_${endDate}.xlsx`)}`);
      } else {
        console.log("已取消操作");
      }
      readline.close();
    }
  );
};

askStartDate();
