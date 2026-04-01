const watchPathInput = document.getElementById("watchPath");
const drivePathInput = document.getElementById("drivePath");
const counter = document.getElementById("counter");
const statusBox = document.getElementById("statusBox");
const sentList = document.getElementById("sentList");

function setStatus(text) {
  statusBox.textContent = text;
}

function formatDate(iso) {
  return new Date(iso).toLocaleString("tr-TR");
}

async function refreshLastSent() {
  const items = await window.bridgeApi.getLastSent();
  sentList.innerHTML = "";
  if (!items.length) {
    const li = document.createElement("li");
    li.textContent = "Henüz gönderim yok.";
    sentList.appendChild(li);
    return;
  }

  for (const item of items) {
    const li = document.createElement("li");
    li.textContent = `${item.name} - ${formatDate(item.sentAt)}`;
    sentList.appendChild(li);
  }
}

async function refreshConfig() {
  const config = await window.bridgeApi.getConfig();
  watchPathInput.value = config.watchPath || "";
  drivePathInput.value = config.drivePath || "";
  counter.textContent = `Kayıtlı gönderim: ${config.sentCount}`;
}

document.getElementById("pickWatch").addEventListener("click", async () => {
  const selected = await window.bridgeApi.chooseWatchFolder();
  if (selected) {
    setStatus("Kaynak klasör güncellendi.");
    await refreshConfig();
  }
});

document.getElementById("pickDrive").addEventListener("click", async () => {
  const selected = await window.bridgeApi.chooseDriveFolder();
  if (selected) {
    setStatus("Hedef klasör güncellendi.");
    await refreshConfig();
  }
});

document.getElementById("sendNow").addEventListener("click", async () => {
  setStatus("Gönderim başlatıldı...");
  const result = await window.bridgeApi.sendNow();
  setStatus(
    `Tamamlandı. Yeni: ${result.sent}, Atlanan: ${result.skipped}, Yoksayılan: ${result.ignored}, Hata: ${result.failed}`
  );
  await refreshConfig();
  await refreshLastSent();
});

window.bridgeApi.onStatus(async (payload) => {
  setStatus(payload.message);
  await refreshConfig();
  await refreshLastSent();
});

async function init() {
  await refreshConfig();
  await refreshLastSent();
}

init();
