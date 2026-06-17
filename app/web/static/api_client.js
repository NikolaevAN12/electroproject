/** Общий клиент API: JSON-ответы и понятные ошибки сети. */
function apiErrorFromPayload(data, status) {
  if (data && data.error) return data.error;
  if (data && Array.isArray(data.detail)) {
    return data.detail.map((d) => d.msg || JSON.stringify(d)).join("; ");
  }
  return `Ошибка сервера (${status})`;
}

async function fetchJson(url, options) {
  if (window.location.protocol === "file:") {
    throw new Error(
      "Страница открыта как файл. Запустите run_web.bat и откройте http://127.0.0.1:8765/"
    );
  }
  let resp;
  try {
    resp = await fetch(url, options);
  } catch (_err) {
    throw new Error(
      "Нет связи с сервером. Запустите run_web.bat и откройте страницу через http://127.0.0.1:8765/ или http://IP-этого-ПК:8765/ (не как файл)."
    );
  }
  const text = await resp.text();
  let data;
  try {
    data = text ? JSON.parse(text) : {};
  } catch (_err) {
    const hint = text.trimStart().startsWith("<")
      ? "Сервер вернул HTML вместо JSON — проверьте, что запущен run_web.py на этом компьютере."
      : (text.slice(0, 200) || "(пустой ответ)");
    throw new Error(hint);
  }
  if (!data.ok) {
    throw new Error(apiErrorFromPayload(data, resp.status));
  }
  return data;
}
