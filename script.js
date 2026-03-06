let sdaAtual = "geral";

async function carregarDados() {

const res = await fetch("/dados?sda=" + sdaAtual);
const dados = await res.json();


// ================= KPI =================

document.getElementById("total").innerText = dados.total;
document.getElementById("postados").innerText = dados.postados;
document.getElementById("progresso").innerText = dados.progresso + "%";

document.querySelector(".barra-fill").style.width =
dados.progresso + "%";


// ================= SDA =================

document.querySelectorAll(".sda-card").forEach(card => {

const nome = card.dataset.sda;

let valor = 0;

if(nome === "geral"){
valor = dados.progresso;
}
else{
valor = dados.sdas[nome] ?? 0;
}

card.querySelector("strong").innerText = valor + "%";
card.querySelector(".mini-bar").style.width = valor + "%";

});


// ================= TABELA =================

const tabela = document.getElementById("tabela");
tabela.innerHTML = "";

dados.tabela.forEach(item => {

const cor =
item.status === "Pendente"
? "status-pendente"
: "status-parcial";

tabela.innerHTML += `
<tr>
<td>${item.item}</td>
<td>${item.setor}</td>
<td>${item.documento}</td>
<td>${item.total}</td>
<td>${item.postados}</td>
<td>${item.comentario}</td>
<td><span class="${cor}">${item.status}</span></td>
</tr>
`;

});

}


// ================= CLICK SDA =================

document.querySelectorAll(".sda-card").forEach(card => {

card.onclick = () => {

document.querySelectorAll(".sda-card")
.forEach(c => c.classList.remove("ativo"));

card.classList.add("ativo");

sdaAtual = card.dataset.sda;

carregarDados();

};

});


// ================= START =================

carregarDados();