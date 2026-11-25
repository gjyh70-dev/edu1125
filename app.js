"use strict";

const form = document.getElementById("trip");
const field = (id) => document.getElementById(id);
const nights = field("nights");
const excelButton = field("save-excel");
let calculatedResult = null;
let savingExcel = false;
const won = (amount) => amount.toLocaleString("ko-KR") + "원";
const dayLength = 86_400_000;
const dailyRate = 25_000;
const lodgingCaps = { seoul: 100_000, metro: 80_000, other: 70_000 };
const remoteMode = () => form.elements.type.value === "remote";
const waterOrAir = () => ["sea", "air"].includes(field("transport").value);

function tripDays() {
  return (field("end").valueAsNumber - field("start").valueAsNumber) / dayLength + 1;
}

function markError(input, message) {
  throw Object.assign(new Error(message), { input });
}

function readNumber(input, defaultValue = 0) {
  const amount = input.value === "" ? defaultValue : input.valueAsNumber;
  if (input.validity.badInput || !Number.isSafeInteger(amount) || amount < 0) {
    markError(input, "금액과 일수는 0 이상의 정수로 입력해 주세요.");
  }
  return amount;
}

function addNight(amount = "", region = "seoul", focus = false) {
  const row = document.createElement("div");
  row.className = "night";
  row.innerHTML = `
    <div class="night-top"><h4></h4><button type="button" class="quiet remove-night">삭제</button></div>
    <div class="fields">
      <label>숙박 지역 <select class="region"><option value="seoul">서울특별시</option><option value="metro">광역시</option><option value="other">그 밖의 지역</option></select></label>
      <label>1박 실제 숙박비 (원)<input class="cost" type="number" min="0" step="1" inputmode="numeric" placeholder="0"></label>
    </div>
    <label class="check extra-night" hidden><input type="checkbox"> 이 숙박의 상한 초과분 지급이 승인됨</label>`;
  row.querySelector(".region").value = region;
  row.querySelector(".cost").value = amount;
  nights.append(row);
  refresh();
  clearResult();
  if (focus) row.querySelector("select").focus();
}

function refresh() {
  const remote = remoteMode();
  field("local").hidden = remote;
  field("remote").hidden = !remote;
  if (!remote) return;

  const grade = field("grade").value;
  const transport = field("transport").value;
  const official = transport === "official";
  field("fare").disabled = official;
  field("halfDays").disabled = official || grade === "actual";
  field("actualCosts").hidden = grade !== "actual";
  field("meal-days-label").hidden = !waterOrAir() || grade === "actual";
  field("inland-label").hidden = !waterOrAir();
  field("mileage-fields").hidden = transport !== "air";
  const fareHints = {
    rail: (grade === "2" ? "일반실" : "특실") + " 기준으로 실제 지급 대상이 되는 금액을 입력하세요.",
    bus: "지급 대상 버스요금 등 실제 운임을 입력하세요.",
    private: "철도·버스 기준 운임을 입력하세요. 불가피한 자가용 이용에는 기관이 인정한 연료비·통행료 등을 입력하세요.",
    air: "좌석 등급 기준과 공적 항공마일리지 사용분을 반영한 지급 대상 운임을 입력하세요.",
    sea: (grade === "2" ? "2등급" : "1등급") + " 기준으로 실제 지급 대상이 되는 금액을 입력하세요.",
    official: "공용 교통수단만 이용한 경우 운임은 0원입니다. 일비는 감액 기준을 적용합니다."
  };
  field("fare-hint").textContent = fareHints[transport];

  const maxNights = tripDays() - 1;
  const count = nights.children.length;
  field("add-night").disabled = !Number.isInteger(maxNights) || count >= maxNights;
  field("night-hint").textContent = maxNights >= 0
    ? "지급 대상 숙박을 1박씩 추가하세요. 입력한 날짜 기준 최대 " + maxNights + "박입니다."
    : "출장일과 복귀일을 먼저 입력하세요.";
  Array.from(nights.children).forEach((row, index) => {
    row.querySelector("h4").textContent = (index + 1) + "박차";
    row.querySelector(".extra-night").hidden = !field("extra").checked || grade !== "2";
    row.querySelector(".remove-night").setAttribute("aria-label", (index + 1) + "박차 숙박 삭제");
  });
}

function calculateLocal(days) {
  const depart = field("depart");
  const arrive = field("arrive");
  if (!depart.value) markError(depart, "출발 시각을 입력해 주세요.");
  if (!arrive.value) markError(arrive, "복귀 시각을 입력해 주세요.");

  const minutes = ((days - 1) * dayLength + arrive.valueAsNumber - depart.valueAsNumber) / 60_000;
  if (minutes <= 0) markError(arrive, "복귀 시각은 출발 시각보다 늦어야 합니다.");
  const base = minutes >= 240 ? 20_000 : 10_000;
  const vehicle = field("localVehicle").value;
  const deduction = { none: 0, official: 10_000, assigned: base }[vehicle];
  const rows = [["출장 정액", base, minutes >= 240 ? "4시간 이상" : "4시간 미만"]];
  if (deduction) rows.push(["차량 조건 감액", -deduction, vehicle === "assigned" ? "전용 차량 배정자" : "공무용 차량 이용 등"]);
  return {
    rows,
    context: "근무지 내 · " + Math.floor(minutes / 60) + "시간 " + minutes % 60 + "분",
    note: days > 1 ? "입력한 구간 전체를 한 번의 연속 출장으로 계산했습니다." : ""
  };
}

function lodgingRow(days, grade) {
  const rows = Array.from(nights.children);
  if (rows.length > days - 1) markError(field("end"), "출장 기간보다 숙박 내역이 많습니다. 날짜를 확인하거나 숙박 내역을 삭제해 주세요.");
  if (waterOrAir() && !field("inland").checked) {
    return ["숙박비", 0, "선박·항공 여행에서 별도 지급 대상 육지 숙박 없음"];
  }

  let amount = 0;
  let limited = false;
  for (const row of rows) {
    const input = row.querySelector(".cost");
    if (!input.value) markError(input, "추가한 숙박의 실제 지출액을 입력해 주세요.");
    const cost = readNumber(input);
    const cap = lodgingCaps[row.querySelector(".region").value];
    const approved = field("extra").checked && row.querySelector(".extra-night input").checked;
    const eligible = grade === "2" ? Math.min(cost, approved ? cap + cap * 0.3 : cap) : cost;
    amount += eligible;
    limited ||= eligible < cost;
  }
  return ["숙박비", amount, rows.length + "박 · " + (limited ? "1박별 상한 적용" : "지급 대상 실비")];
}

function daysWithin(input, days, defaultValue) {
  const value = readNumber(input, defaultValue);
  if (value > days) markError(input, "지급 대상 일수는 전체 출장일수를 넘을 수 없습니다.");
  return value;
}

function calculateRemote(days) {
  const grade = field("grade").value;
  const actual = grade === "actual";
  const official = field("transport").value === "official";
  const fare = official ? 0 : readNumber(field("fare"));
  const half = actual ? 0 : official ? days : daysWithin(field("halfDays"), days, 0);
  const meals = !actual && waterOrAir() ? daysWithin(field("mealDays"), days, 0) : days;
  const daily = actual ? readNumber(field("actualDaily")) : dailyRate * days - dailyRate / 2 * half;
  const mealCost = actual ? readNumber(field("actualMeal")) : dailyRate * meals;
  const rows = [
    ["운임", fare, official ? "공용 교통수단 이용" : "지급 대상 왕복 운임"],
    ["일비", daily, actual ? "실비" : days + "일" + (half ? " 중 " + half + "일 절반 지급" : " × " + won(dailyRate))],
    lodgingRow(days, grade),
    ["식비", mealCost, actual ? "지급 대상 실비" : meals + "일 × " + won(dailyRate)]
  ];
  const bonus = field("transport").value === "air" ? readNumber(field("mileage")) : 0;
  if (bonus) rows.push(["공적 항공마일리지 추가 일비", bonus, "기관 인정액"]);
  let note = fare || rows[2][1] ? "운임·숙박비는 출장 종료 다음 날부터 2주 이내에 증빙을 갖춰 정산하세요." : "";
  if (field("extra").checked && nights.querySelector(".extra-night input:checked") && grade === "2") {
    note = "숙박비 상한 초과분은 출장 종료 다음 날부터 1주 이내에 증빙을 갖춰 정산하세요. " + note;
  }
  return { rows, context: "근무지 외 · " + days + "일", note };
}

function clearResult(message = "출장 정보를 입력하고 계산해 주세요.") {
  calculatedResult = null;
  excelButton.disabled = true;
  field("error").hidden = true;
  field("total").textContent = "—";
  field("result-context").textContent = message;
  field("result-dates").hidden = true;
  field("breakdown").replaceChildren();
  field("result-note").hidden = true;
  field("print-result").hidden = true;
}

function showResult({ rows, context, note }) {
  const total = rows.reduce((sum, [, amount]) => sum + amount, 0);
  if (!Number.isSafeInteger(total)) markError(null, "표시 가능한 금액 범위를 초과했습니다.");
  const start = field("start").value;
  const end = field("end").value;
  calculatedResult = { start, end, rows, context, note, total };
  excelButton.disabled = savingExcel;
  field("total").textContent = won(total);
  field("result-context").textContent = context;
  field("result-dates").textContent = "출장일 " + start + " · 복귀일 " + end;
  field("result-dates").hidden = false;
  for (const [label, amount, detail] of rows) {
    const row = document.createElement("div");
    const name = document.createElement("dt");
    const description = document.createElement("small");
    const value = document.createElement("dd");
    name.textContent = label;
    description.textContent = detail;
    value.textContent = won(amount);
    name.append(description);
    row.append(name, value);
    field("breakdown").append(row);
  }
  field("result-note").textContent = note;
  field("result-note").hidden = !note;
  field("print-result").hidden = false;
}

form.addEventListener("submit", (event) => {
  event.preventDefault();
  clearResult();
  try {
    if (!field("start").value) markError(field("start"), "출장일을 입력해 주세요.");
    if (!field("end").value) markError(field("end"), "복귀일을 입력해 주세요.");
    const days = tripDays();
    if (days < 1) markError(field("end"), "복귀일은 출장일과 같거나 그 이후여야 합니다.");
    showResult(remoteMode() ? calculateRemote(days) : calculateLocal(days));
  } catch (error) {
    if (!Object.hasOwn(error, "input")) throw error;
    field("error").textContent = error.message;
    field("error").hidden = false;
    const details = error.input?.closest("details");
    if (details) details.open = true;
    error.input?.focus();
  }
});

form.addEventListener("input", () => {
  if (calculatedResult) clearResult("입력값이 변경되었습니다. 여비를 다시 계산해 주세요.");
  field("error").hidden = true;
  refresh();
});
form.addEventListener("reset", () => {
  nights.replaceChildren();
  field("adjustments").open = false;
  clearResult();
  queueMicrotask(refresh);
});
field("add-night").addEventListener("click", () => addNight("", "seoul", true));
field("print-result").addEventListener("click", () => {
  if (!field("print-result").hidden) window.print();
});
excelButton.addEventListener("click", async () => {
  if (!calculatedResult || savingExcel) return;
  savingExcel = true;
  excelButton.disabled = true;
  try {
    await downloadTravelWorkbook(calculatedResult);
  } catch (error) {
    console.error(error);
    field("error").textContent = "엑셀 파일을 저장하지 못했습니다. 다시 시도해 주세요.";
    field("error").hidden = false;
  } finally {
    savingExcel = false;
    excelButton.disabled = !calculatedResult;
  }
});
nights.addEventListener("click", (event) => {
  if (!event.target.matches(".remove-night")) return;
  event.target.closest(".night").remove();
  refresh();
  clearResult();
  field("add-night").focus();
});
field("example").addEventListener("click", () => {
  const type = form.elements.type.value;
  form.reset();
  form.elements.type.value = type;
  field("start").value = "2026-11-25";
  field("end").value = remoteMode() ? "2026-11-26" : "2026-11-25";
  if (remoteMode()) {
    field("fare").value = "60000";
    addNight("75000", "metro");
  } else {
    field("depart").value = "09:00";
    field("arrive").value = "14:00";
  }
  refresh();
  form.requestSubmit();
});

refresh();
