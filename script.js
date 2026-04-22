const API_URL_INPUT = document.querySelector("input");

let API_URL = localStorage.getItem("api_url") || "";

// salvar URL
function salvarURL() {
  const url = API_URL_INPUT.value.trim();
  if (!url.includes("/exec")) {
    alert("URL inválida");
    return;
  }

  localStorage.setItem("api_url", url);
  API_URL = url;

  alert("URL salva com sucesso!");
}

// puxar dados
async function puxarDados() {
  try {
    const res = await fetch(API_URL);
    const data = await res.json();

    console.log("Dados recebidos:", data);

    alert("Conectado com sucesso!");
  } catch (err) {
    console.error(err);
    alert("Erro ao conectar com Apps Script");
  }
}

// enviar dados
async function enviarDados(dados) {
  try {
    const res = await fetch(API_URL, {
      method: "POST",
      body: JSON.stringify(dados),
    });

    const data = await res.json();

    console.log("Enviado:", data);

    alert("Dados enviados com sucesso!");
  } catch (err) {
    console.error(err);
    alert("Erro ao enviar dados");
  }
}

// eventos dos botões
document.addEventListener("DOMContentLoaded", () => {
  const btnSalvar = document.querySelector("button");
  const btnPuxar = document.querySelectorAll("button")[1];
  const btnEnviar = document.querySelectorAll("button")[2];

  if (btnSalvar) btnSalvar.onclick = salvarURL;
  if (btnPuxar) btnPuxar.onclick = puxarDados;
  if (btnEnviar) btnEnviar.onclick = () => {
    enviarDados({ teste: "ok" });
  };

  if (API_URL) {
    API_URL_INPUT.value = API_URL;
  }
});
