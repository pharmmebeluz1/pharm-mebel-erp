const translations = {
  uz: {
    tagline: "MEBEL ISHLAB CHIQARISH BOSHQARUV TIZIMI",
    tashkent: "Toshkent", welcome: "dasturiga xush kelibsiz!", opening: "Dastur ochilmoqda...", enter: "Kirish",
    loginTitle: "Kirish", loginSubtitle: "Hisobingizga kiring", username: "Foydalanuvchi nomi", password: "Parol",
    remember: "Meni eslab qol", or: "yoki", smsLogin: "SMS orqali kirish", chooseRole: "Rolni tanlang",
    chooseRoleSub: "Iltimos, o‘zingizning rolingizni tanlang", admin: "Admin", management: "Boshqaruv",
    worker: "Ishchi", workshopEmployee: "Korxona xodimi", driver: "Shofyor", delivery: "Yetkazib berish",
    manager: "Menejer", clients: "Mijozlar bilan ish", constructor: "Konstruktor", designing: "Loyihalash",
    customer: "Mijoz", myContract: "Mening shartnomam", fillFields: "Login va parolni kiriting.",
    success: "Kirish muvaffaqiyatli.", smsSoon: "SMS orqali kirish keyingi bosqichda ulanadi.", roleSelected: "bo‘limi tanlandi.",
    fullContract: "To‘liq shartnoma", loadingContract: "Shartnoma ochilmoqda...", contractNumber: "Shartnoma raqami",
    orderCode: "Buyurtma kodi", contractStatus: "Holati", acceptContractText: "Shartnomaning barcha bandlarini o‘qidim va roziman.",
    printPdf: "Chop etish / PDF", confirmContract: "Shartnomani tasdiqlash", contractSavedNote: "Tasdiq sanasi va vaqti tizimda saqlanadi.",
    contractLoaded: "Shartnoma avtomatik ochildi.", contractConfirmed: "Shartnoma tasdiqlandi.", acceptFirst: "Avval shartnoma bilan tanishganingizni belgilang.",
    confirmed: "Tasdiqlangan", notConfirmed: "Tasdiqlanmagan"
  },
  ru: {
    tagline: "СИСТЕМА УПРАВЛЕНИЯ МЕБЕЛЬНЫМ ПРОИЗВОДСТВОМ", tashkent: "Ташкент", welcome: "Добро пожаловать в программу!",
    opening: "Программа запускается...", enter: "Войти", loginTitle: "Вход", loginSubtitle: "Войдите в свою учётную запись",
    username: "Имя пользователя", password: "Пароль", remember: "Запомнить меня", or: "или", smsLogin: "Войти через SMS",
    chooseRole: "Выберите роль", chooseRoleSub: "Пожалуйста, выберите свою роль", admin: "Админ", management: "Управление",
    worker: "Рабочий", workshopEmployee: "Сотрудник предприятия", driver: "Водитель", delivery: "Доставка", manager: "Менеджер",
    clients: "Работа с клиентами", constructor: "Конструктор", designing: "Проектирование", customer: "Клиент", myContract: "Мой договор",
    fillFields: "Введите логин и пароль.", success: "Вход выполнен успешно.", smsSoon: "Вход через SMS будет подключён на следующем этапе.",
    roleSelected: "— роль выбрана.", fullContract: "Полный договор", loadingContract: "Договор открывается...", contractNumber: "Номер договора",
    orderCode: "Код заказа", contractStatus: "Статус", acceptContractText: "Я прочитал(а) все пункты договора и согласен(на).",
    printPdf: "Печать / PDF", confirmContract: "Подтвердить договор", contractSavedNote: "Дата и время подтверждения сохраняются в системе.",
    contractLoaded: "Договор открыт автоматически.", contractConfirmed: "Договор подтверждён.", acceptFirst: "Сначала отметьте, что вы ознакомились с договором.",
    confirmed: "Подтверждено", notConfirmed: "Не подтверждено"
  },
  en: {
    tagline: "FURNITURE PRODUCTION MANAGEMENT SYSTEM", tashkent: "Tashkent", welcome: "Welcome to the application!", opening: "The app is starting...",
    enter: "Sign in", loginTitle: "Sign in", loginSubtitle: "Access your account", username: "Username", password: "Password", remember: "Remember me",
    or: "or", smsLogin: "Sign in with SMS", chooseRole: "Choose a role", chooseRoleSub: "Please select your role", admin: "Admin", management: "Management",
    worker: "Worker", workshopEmployee: "Company employee", driver: "Driver", delivery: "Delivery", manager: "Manager", clients: "Client relations",
    constructor: "Designer", designing: "Design", customer: "Customer", myContract: "My contract", fillFields: "Enter your username and password.",
    success: "Signed in successfully.", smsSoon: "SMS sign-in will be connected in the next stage.", roleSelected: "role selected.",
    fullContract: "Full contract", loadingContract: "Opening the contract...", contractNumber: "Contract number", orderCode: "Order code", contractStatus: "Status",
    acceptContractText: "I have read and agree to all terms of the contract.", printPdf: "Print / PDF", confirmContract: "Confirm contract",
    contractSavedNote: "The confirmation date and time are stored in the system.", contractLoaded: "The contract opened automatically.",
    contractConfirmed: "Contract confirmed.", acceptFirst: "First confirm that you have read the contract.", confirmed: "Confirmed", notConfirmed: "Not confirmed"
  }
};

const languageNames = [
  { code: "uz", label: "O‘zbekcha" }, { code: "ru", label: "Русский" }, { code: "en", label: "English" }
];

let currentLanguage = localStorage.getItem("mebel360-language") || "uz";
let toastTimer;
let activeContract = null;

function createLanguageBars() {
  document.querySelectorAll("[data-language-bar]").forEach((bar) => {
    bar.innerHTML = languageNames.map(({ code, label }) =>
      `<button type="button" class="language-button" data-language="${code}">${label}</button>`
    ).join("");
  });
}

function applyLanguage(language) {
  currentLanguage = translations[language] ? language : "uz";
  localStorage.setItem("mebel360-language", currentLanguage);
  document.documentElement.lang = currentLanguage;
  document.querySelectorAll("[data-i18n]").forEach((element) => {
    const key = element.dataset.i18n;
    if (translations[currentLanguage][key]) element.textContent = translations[currentLanguage][key];
  });
  document.querySelectorAll("[data-i18n-placeholder]").forEach((element) => {
    const key = element.dataset.i18nPlaceholder;
    if (translations[currentLanguage][key]) element.placeholder = translations[currentLanguage][key];
  });
  document.querySelectorAll("[data-language]").forEach((button) => {
    button.classList.toggle("active", button.dataset.language === currentLanguage);
  });
  updateTashkentClock();
  if (activeContract) updateContractConfirmationUI(activeContract);
}

function showScreen(screenId) {
  document.querySelectorAll(".screen").forEach((screen) => screen.classList.remove("active"));
  const target = document.getElementById(screenId);
  if (target) {
    target.classList.add("active");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }
}

function updateTashkentClock() {
  const now = new Date();
  const locale = currentLanguage === "ru" ? "ru-RU" : currentLanguage === "en" ? "en-GB" : "uz-UZ";
  const time = new Intl.DateTimeFormat("en-GB", { timeZone: "Asia/Tashkent", hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false }).format(now);
  const weekday = new Intl.DateTimeFormat(locale, { timeZone: "Asia/Tashkent", weekday: "long" }).format(now);
  const date = new Intl.DateTimeFormat(locale, { timeZone: "Asia/Tashkent", day: "2-digit", month: "long", year: "numeric" }).format(now);
  document.getElementById("liveClock").textContent = time;
  document.getElementById("weekdayText").textContent = weekday.charAt(0).toUpperCase() + weekday.slice(1);
  document.getElementById("dateText").textContent = date;
}

function runProgress() {
  let progress = 0;
  const fill = document.getElementById("progressFill");
  const text = document.getElementById("progressText");
  const timer = setInterval(() => {
    const jump = progress < 55 ? 4 : progress < 85 ? 2 : 1;
    progress = Math.min(progress + jump, 100);
    fill.style.width = `${progress}%`;
    text.textContent = `${progress}%`;
    if (progress >= 100) clearInterval(timer);
  }, 70);
}

function showToast(message) {
  const toast = document.getElementById("toast");
  toast.textContent = message;
  toast.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove("show"), 2800);
}

function setTheme(dark) {
  document.body.classList.toggle("dark", dark);
  localStorage.setItem("mebel360-theme", dark ? "dark" : "light");
  document.querySelectorAll("#themeToggle, .theme-clone").forEach((button) => { button.textContent = dark ? "☾" : "☀"; });
}

async function submitLogin(event) {
  event.preventDefault();
  const username = document.getElementById("username").value.trim();
  const password = document.getElementById("password").value.trim();
  if (!username || !password) { showToast(translations[currentLanguage].fillFields); return; }
  const button = event.submitter;
  if (button) button.disabled = true;
  try {
    const response = await fetch("/api/login", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ username, password, remember: document.getElementById("rememberMe").checked })
    });
    const result = await response.json();
    if (!response.ok || !result.ok) throw new Error(result.message || "Login failed");
    showToast(translations[currentLanguage].success);
    if (result.auto_contract || result.next_screen === "contractScreen") {
      setTimeout(() => { showScreen("contractScreen"); loadCustomerContract(); }, 420);
    } else {
      setTimeout(() => showScreen("roleScreen"), 420);
    }
  } catch (error) { showToast(error.message || translations[currentLanguage].fillFields); }
  finally { if (button) button.disabled = false; }
}

function valueAtPath(object, path) {
  return path.split(".").reduce((current, key) => current && current[key], object);
}

function fillContract(contractData) {
  activeContract = contractData;
  document.querySelectorAll("[data-contract]").forEach((element) => {
    const value = valueAtPath(contractData, element.dataset.contract);
    element.textContent = value ?? "—";
  });
  document.getElementById("contractHeaderNumber").textContent = contractData.contract.number;
  document.getElementById("summaryContractNumber").textContent = contractData.contract.number;
  document.getElementById("summaryOrderCode").textContent = contractData.order.code;
  document.getElementById("summaryContractStatus").textContent = contractData.contract.status;
  updateContractConfirmationUI(contractData);
}

function updateContractConfirmationUI(contractData) {
  const approved = Boolean(contractData.contract.confirmed_at);
  const status = document.getElementById("summaryContractStatus");
  const approval = document.getElementById("contractElectronicApproval");
  const checkbox = document.getElementById("contractAcceptCheckbox");
  const confirmButton = document.getElementById("confirmContractBtn");
  const acceptRow = document.getElementById("contractAcceptRow");
  if (approved) {
    const text = `${translations[currentLanguage].confirmed}: ${contractData.contract.confirmed_date || "—"}`;
    approval.textContent = text;
    status.textContent = text;
    status.classList.add("confirmed");
    checkbox.checked = true;
    checkbox.disabled = true;
    confirmButton.disabled = true;
    confirmButton.querySelector("[data-i18n]").textContent = translations[currentLanguage].confirmed;
    acceptRow.classList.add("accepted");
  } else {
    approval.textContent = translations[currentLanguage].notConfirmed;
    status.textContent = contractData.contract.status;
    status.classList.remove("confirmed");
    checkbox.disabled = false;
    confirmButton.disabled = false;
    confirmButton.querySelector("[data-i18n]").textContent = translations[currentLanguage].confirmContract;
    acceptRow.classList.remove("accepted");
  }
}

async function loadCustomerContract() {
  const loading = document.getElementById("contractLoading");
  const content = document.getElementById("contractContent");
  loading.hidden = false;
  content.hidden = true;
  try {
    const response = await fetch("/api/customer/contract", { headers: { "Accept": "application/json" } });
    const result = await response.json();
    if (!response.ok || !result.ok) throw new Error(result.message || "Contract error");
    fillContract(result.data);
    loading.hidden = true;
    content.hidden = false;
    showToast(translations[currentLanguage].contractLoaded);
  } catch (error) {
    loading.innerHTML = `<strong>${error.message}</strong>`;
    showToast(error.message);
  }
}

async function confirmContract() {
  if (!activeContract) return;
  const checkbox = document.getElementById("contractAcceptCheckbox");
  if (!checkbox.checked) { showToast(translations[currentLanguage].acceptFirst); return; }
  const button = document.getElementById("confirmContractBtn");
  button.disabled = true;
  try {
    const response = await fetch("/api/customer/contract/confirm", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ contract_id: activeContract.contract.id, accepted: true })
    });
    const result = await response.json();
    if (!response.ok || !result.ok) throw new Error(result.message || "Confirm error");
    activeContract.contract.confirmed_at = result.confirmed_at;
    activeContract.contract.confirmed_date = result.confirmed_date;
    activeContract.contract.status = "Mijoz tasdiqladi";
    updateContractConfirmationUI(activeContract);
    showToast(translations[currentLanguage].contractConfirmed);
  } catch (error) {
    button.disabled = false;
    showToast(error.message);
  }
}

async function selectRole(role, card) {
  document.querySelectorAll(".role-card").forEach((item) => item.classList.remove("featured"));
  card.classList.add("featured");
  try {
    const response = await fetch("/api/select-role", {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ role })
    });
    const result = await response.json();
    if (!response.ok || !result.ok) throw new Error(result.message || "Role error");
    const roleName = card.querySelector("strong").textContent;
    showToast(`${roleName} ${translations[currentLanguage].roleSelected}`);
    if (role === "customer") {
      setTimeout(() => { showScreen("contractScreen"); loadCustomerContract(); }, 320);
    }
  } catch (error) { showToast(error.message); }
}

createLanguageBars();
applyLanguage(currentLanguage);
setTheme(localStorage.getItem("mebel360-theme") === "dark");
updateTashkentClock();
setInterval(updateTashkentClock, 1000);
runProgress();

document.addEventListener("click", (event) => {
  const languageButton = event.target.closest("[data-language]");
  if (languageButton) applyLanguage(languageButton.dataset.language);
  const backButton = event.target.closest("[data-back]");
  if (backButton) showScreen(backButton.dataset.back);
  const roleCard = event.target.closest("[data-role]");
  if (roleCard) selectRole(roleCard.dataset.role, roleCard);
  if (event.target.closest("#themeToggle, .theme-clone")) setTheme(!document.body.classList.contains("dark"));
});

document.getElementById("openLoginBtn").addEventListener("click", () => showScreen("loginScreen"));
document.getElementById("loginForm").addEventListener("submit", submitLogin);
document.getElementById("smsBtn").addEventListener("click", () => showToast(translations[currentLanguage].smsSoon));
document.getElementById("passwordToggle").addEventListener("click", () => {
  const input = document.getElementById("password");
  input.type = input.type === "password" ? "text" : "password";
});
document.getElementById("printContractBtn").addEventListener("click", () => window.print());
document.getElementById("confirmContractBtn").addEventListener("click", confirmContract);

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => navigator.serviceWorker.register("/static/sw.js").catch(() => {}));
}
