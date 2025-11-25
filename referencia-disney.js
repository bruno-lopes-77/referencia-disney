let tabelas;
let tabelasFiltradas;
let abaExcel;

const nomesColunasNomes = [
	"nome",
	"nome-completo",
	"nome-descritivo"
];
const nomesColunasNomesOriginais = [
	"nome-original",
	"nome-original-completo",
	"nome-original-descritivo"
];
const nomesColunasNomesInglês = [
	"nome-inglês",
	"nome-inglês-completo",
	"nome-inglês-descritivo"
];
const nomesColunasNomesGeral = [
	["nomes", nomesColunasNomes],
	["nomes-originais", nomesColunasNomesOriginais],
	["nomes-inglês", nomesColunasNomesInglês]
];
const nomesColunasOutrosNomes = [
	"outros-nomes",
	"outros-nomes-originais",
	"outros-nomes-inglês"
];
const nomesColunasListasReferências = [
	"origem",
	"espécie",
	"amores-desamores",
	"identidades-secretas",
	"alter-egos",
	"animais-estimação",
	"objetos",
	"carro",
	"lugares",
	"universo",
	"hierarquia-superior",
	"hierarquia-inferior"
];
const nomesColunasHtml = [
	"número",
	"nomes",
	"outros-nomes",
	"nomes-originais",
	"outros-nomes-originais",
	"língua-original",
	"estreia",
	"origem",
	"universo",
	"lugares",
	"género",
	"espécie",
	"amores-desamores",
	"alter-egos",
	"identidades-secretas",
	"animais-estimação",
	"objetos",
	"carro",
	"hierarquia-superior",
	"hierarquia-inferior",
];

const nomesColunasHtmlOcultas = {
	"personagens": [],
	"origens": ["língua-original", "nomes-originais", "outros-nomes-originais", "estreia", "género", "espécie", "amores-desamores", "alter-egos",
					"identidades-secretas", "animais-estimação", "hierarquia-superior", "hierarquia-inferior"],
	"espécies": ["estreia", "origem", "género", "amores-desamores", "alter-egos", "identidades-secretas", "animais-estimação", "objetos", "carro", "lugares",
					"universo"],
	"objetos": ["género", "espécie", "amores-desamores", "alter-egos", "identidades-secretas", "animais-estimação", "carro", "hierarquia-superior",
					"hierarquia-inferior"],
	"carros": ["género", "espécie", "amores-desamores", "alter-egos", "animais-estimação", "objetos", "hierarquia-superior", "hierarquia-inferior"],
	"lugares": ["género", "espécie", "amores-desamores", "alter-egos", "identidades-secretas", "animais-estimação"],
	"universos": ["género", "espécie", "amores-desamores", "identidades-secretas", "animais-estimação", "hierarquia-superior", "hierarquia-inferior"],
};
const matrizReferênciasCruzadas = {
	"personagens": {
		"origem": ["origens", "origem"],
		"espécie": ["espécies", "espécie"],
		"amores-desamores": ["personagens", "amores-desamores"],
		"identidades-secretas": ["personagens", "identidades-secretas"],
		"alter-egos": ["personagens", "alter-egos"],
		"animais-estimação": ["personagens", "animais-estimação"],
		"objetos": ["objetos"],
		"carro": ["carros"],
		"lugares": ["lugares", "lugares"],
		"universo": ["universos", "universo"],
		"hierarquia-superior": ["personagens"],
		"hierarquia-inferior": ["personagens", "hierarquia-superior"],
	},
	"origens": {
		"origem": ["personagens"],
		"espécie": ["espécies"],
		"amores-desamores": ["origens", "amores-desamores"],
		"identidades-secretas": ["origens", "identidades-secretas"],
		"alter-egos": ["origens", "alter-egos"],
		"animais-estimação": ["origens", "animais-estimação"],
		"objetos": ["objetos"],
		"carro": ["carros"],
		"lugares": ["lugares"],
		"universo": ["universos"],
		"hierarquia-superior": ["origens", "hierarquia-inferior"],
		"hierarquia-inferior": ["origens"],
	},
	"espécies": {
		"origem": ["origens", "espécie"],
		"espécie": ["personagens"],
		"amores-desamores": ["espécies", "amores-desamores"],
		"identidades-secretas": ["espécies", "identidades-secretas"],
		"alter-egos": ["espécies", "alter-egos"],
		"animais-estimação": ["espécies", "animais-estimação"],
		"objetos": ["objetos"],
		"carro": ["carros"],
		"lugares": ["lugares", "espécie"],
		"universo": ["universos", "espécie"],
		"hierarquia-superior": ["espécies", "hierarquia-inferior"],
		"hierarquia-inferior": ["espécies"],
	},
	"objetos": {
		"origem": ["origens", "objetos"],
		"espécie": ["espécies", "objetos"],
		"amores-desamores": ["objetos", "amores-desamores"],
		"identidades-secretas": ["objetos", "identidades-secretas"],
		"alter-egos": ["objetos", "alter-egos"],
		"animais-estimação": ["objetos", "animais-estimação"],
		"objetos": ["personagens", "objetos"],
		"carro": ["carros", "objetos"],
		"lugares": ["lugares", "objetos"],
		"universo": ["universos", "objetos"],
		"hierarquia-superior": ["objetos", "hierarquia-inferior"],
		"hierarquia-inferior": ["objetos"],
	},
	"carros": {
		"origem": ["origens", "carro"],
		"espécie": ["espécies", "carro"],
		"amores-desamores": ["carros", "amores-desamores"],
		"identidades-secretas": ["carros", "identidades-secretas"],
		"alter-egos": ["carros", "alter-egos"],
		"animais-estimação": ["carros", "animais-estimação"],
		"objetos": ["objetos"],
		"carro": ["personagens", "carro"],
		"lugares": ["lugares", "carro"],
		"universo": ["universos", "carro"],
		"hierarquia-superior": ["carros", "hierarquia-inferior"],
		"hierarquia-inferior": ["carros"],
	},
	"lugares": {
		"origem": ["origens", "lugares"],
		"espécie": ["espécies"],
		"amores-desamores": ["lugares", "amores-desamores"],
		"identidades-secretas": ["lugares", "identidades-secretas"],
		"alter-egos": ["lugares", "alter-egos"],
		"animais-estimação": ["lugares", "animais-estimação"],
		"objetos": ["objetos"],
		"carro": ["carros"],
		"lugares": ["personagens"],
		"universo": ["universos", "lugares"],
		"hierarquia-superior": ["lugares", "hierarquia-inferior"],
		"hierarquia-inferior": ["lugares"],
	},
	"universos": {
		"origem": ["origens", "universo"],
		"espécie": ["espécies"],
		"amores-desamores": ["universos", "amores-desamores"],
		"identidades-secretas": ["universos", "identidades-secretas"],
		"alter-egos": ["universos", "alter-egos"],
		"animais-estimação": ["universos", "animais-estimação"],
		"objetos": ["objetos"],
		"carro": ["carros"],
		"lugares": ["lugares"],
		"universo": ["personagens"],
		"hierarquia-superior": ["universos", "hierarquia-inferior"],
		"hierarquia-inferior": ["universos"],
	},
};

String.prototype.ordem = function () {
	let res = "";
	switch (this.toString()) {
		case "1":
			res = "I";
			break;
		case "2":
			res = "II";
			break;
		case "3":
			res = "III";
			break;
		case "4":
			res = "IV";
			break;
		case "5":
			res = "V";
			break;
		default:
			res = this.toString();
			break;
	}
	return res;
};

String.prototype.língua = function () {
	let res = "";
	switch (this.toString()) {
		case "en":
			res = "Inglês";
			break;
		case "it":
			res = "Italiano";
			break;
		case "pt":
			res = "Português";
			break;
		default:
			res = this.toString();
	}
	return res;
};

String.prototype.listaOutrosNomes = function (tab) {
	let res = [];
	const listaItens = this.toString().split(/\s*,\s*/).filter(e => e);
	for (const item of listaItens) {
		const a = item.split("/");
		if (a[1]) {
			a[1] = tabelas[tab][a[1]] ? tab + "-" + a[1]: undefined;
		}
		res.push(a);
	}
	return res;
};

String.prototype.listaItens = function () {
	return this.toString().split(/\s*,\s*/).filter(e => e);
};

String.prototype.formatarNome = function () {
	return this.toString()	.replace(/`/gu,  "<i>")
							.replace(/´/gu,  "</i>")
							.replace(/\[/gu, "<cite>")
							.replace(/\]/gu, "</cite>")
							.replace(/\{/gu, "<span class=trocadilho>")
							.replace(/\}/gu, "</span>");
};

String.prototype.prepararNomeOrdenação = function () {
	return this.toString()	.replace(/`/gu,  "")
							.replace(/´/gu,  "")
							.replace(/\[/gu, "")
							.replace(/\]/gu, "")
							.replace(/\{/gu, "")
							.replace(/\}/gu, "");
};

Array.prototype.processarReferênciasCruzadas = function (tab) {
	let res = [];
	for (cód of this) {
		res.push([
			tabelas[tab][tab + "-" + cód]?.["nome-padrão"] ?? cód,
			tabelas[tab][tab + "-" + cód] ? tab + "-" + cód : undefined
		]);
	}
	return res;
};

Array.prototype.formatarNomesHtml = function (ordem) {
	let res = "";
	let nome = this[0];
	let nomeCompleto = this[1];
	let nomeDescritivo = this[2];
	let n = nome ?? nomeCompleto ?? nomeDescritivo ?? "";
	if (n) {
		res += "<span class=nome-principal" + (ordem ? " data-ordem=\"" + ordem + "\"" : "") + ">" + n.formatarNome() + "</span>";
	}
	if (nome && nomeCompleto) {
		res += "<br><span class=nome-completo>" + nomeCompleto.formatarNome() + "</span>";
	}
	if ((nome || nomeCompleto) && nomeDescritivo) {
		res += "<br><span class=nome-descritivo>(" + nomeDescritivo.formatarNome() + ")</span>";
	}
	return res;
};

Array.prototype.formatarListaNomesHtml = function (tab) {
	let a = [];
	for (const i of this) {
		a.push(i.formatarNomeHtml(tab));
	}
	return a.join(", ");
};

Array.prototype.formatarNomeHtml = function (tab) {
	let ret = "";
	let nome = this[0].formatarNome();
	if (this[1]) {
		ret += "<a href=\"#" + this[1] + "\">" + nome + "</a>";
	}
	else {
		ret += nome;
	}
	return ret;
};

async function processarDados() {
	const respostaFicheiro = await fetch('referencia-disney.xlsx');
	const dadosFicheiro = await respostaFicheiro.blob();
	const leitor = new FileReader();
	leitor.onload = function (e) {
		const arrayFicheiro = new Uint8Array(e.target.result);
		const dadosExcel = XLSX.read(arrayFicheiro, {type: 'array'});
		const nomesAbas = dadosExcel.SheetNames;
		abasExcel = {};
		for (const nomeAba of nomesAbas) {
			const aba = dadosExcel.Sheets[nomeAba];
			const conteúdoAba = XLSX.utils.sheet_to_json(aba);
			abasExcel[nomeAba] = conteúdoAba;
		}
		construirObjetoTabelas();
		formatarTabelasHtml();
		preencherTabelaHtml();
	};
	leitor.readAsArrayBuffer(dadosFicheiro);
};

function construirObjetoTabelas() {
	// Construção do objeto das tabelas
	tabelas = {};
	for (const nomeTabela in abasExcel) {
		tabelas[nomeTabela] = {};
		for (const linhaExcel of abasExcel[nomeTabela]) {
			// Criação de listas de nomes
			for (const i of nomesColunasNomesGeral) {
				linhaExcel[i[0]] = listaNomes(linhaExcel, i[1]);
			}
			let nomePadrão = linhaExcel["nomes"][0] ?? linhaExcel["nomes"][1] ?? linhaExcel["nomes"][2];
			if (!nomePadrão) {
				nomePadrão = linhaExcel["nomes-originais"][0] ?? linhaExcel["nomes-originais"][1] ?? linhaExcel["nomes-originais"][2];
				if (nomePadrão) {
					nomePadrão = "`" + nomePadrão + "´";
				}
				else {
					nomePadrão = linhaExcel["nomes-inglês"][0] ?? linhaExcel["nomes-inglês"][1] ?? linhaExcel["nomes-inglês"][2];
					if (nomePadrão) {
						nomePadrão = "`" + nomePadrão + "´";
					}
					else {
						nomePadrão = "";
					}
				}
			}
			linhaExcel["nome-padrão"] = nomePadrão;
		}
		abasExcel[nomeTabela].sort(comparadorLinhas);
		for (const linhaExcel of abasExcel[nomeTabela]) {
			const id = nomeTabela + "-" + linhaExcel.código;
			tabelas[nomeTabela][id] = linhaExcel;
			// Listas: colocar strings como arrays
			for (const nomeColuna of nomesColunasOutrosNomes) {
				linhaExcel[nomeColuna] = (linhaExcel[nomeColuna] ?? "").toString().listaOutrosNomes(nomeTabela);
				linhaExcel[nomeColuna].sort(comparadorNomes);
			}
			linhaExcel.estreia = (linhaExcel.estreia ?? "").toString().listaItens();
			for (const nomeColuna of nomesColunasListasReferências) {
				linhaExcel[nomeColuna] = (linhaExcel[nomeColuna] ?? "").toString().listaItens();
			}
		}
	}
	// Filtragem de linhas para referências
	tabelasFiltradas = {};
	for (const nomeTabela in tabelas) {
		tabelasFiltradas[nomeTabela] = {};
		for (const nomeColuna of nomesColunasListasReferências) {
			tabelasFiltradas[nomeTabela][nomeColuna] = [];
			for (const id in tabelas[nomeTabela]) {
				if (tabelas[nomeTabela][id][nomeColuna].length > 0) {
					tabelasFiltradas[nomeTabela][nomeColuna].push(tabelas[nomeTabela][id]);
					switch (nomeColuna) {
						case "identidades-secretas":
							tabelas[nomeTabela][id]["identidade-secreta"] = true;
							break;
						case "alter-egos":
							tabelas[nomeTabela][id]["alter-ego"] = true;
							break;
						case "animais-estimação":
							tabelas[nomeTabela][id]["animal-estimação"] = true;
							break;
						case "hierarquia-inferior":
							if (nomeTabela === "personagens") {
								tabelas[nomeTabela][id]["grupo-personagens"] = true;
							}
					}
				}
			}
		}
	}
	// Preenchimento de referências cruzadas
	for (const nomeTabela in tabelasFiltradas) {
		for (const nomeColuna of nomesColunasListasReferências) {
			for (const linha of tabelasFiltradas[nomeTabela][nomeColuna]) {
				preencherReferênciasCruzadas(linha, nomeColuna, matrizReferênciasCruzadas[nomeTabela][nomeColuna]);
			}
		}
	}
	// Processamento de referências cruzadas
	for (const nomeTabela in tabelas) {
		for (const id in tabelas[nomeTabela]) {
			for (const nomeColuna of nomesColunasListasReferências) {
				tabelas[nomeTabela][id][nomeColuna + "-nomes"] =
					tabelas[nomeTabela][id][nomeColuna].processarReferênciasCruzadas(matrizReferênciasCruzadas[nomeTabela][nomeColuna][0]);
				tabelas[nomeTabela][id][nomeColuna + "-nomes"].sort(comparadorNomes);
			}
		}
	}
	console.log(tabelas);
}

function formatarTabelasHtml() {
	for (const nomeTabela in tabelas) {
		for (const id in tabelas[nomeTabela]) {
			linha = tabelas[nomeTabela][id];
			// Formatação de campos simples
			linha["html-ordem"] = (linha.ordem ?? "").toString().ordem();
			linha["html-língua-original"] = (linha["língua-original"] ?? "").língua();
			linha["html-estreia"] = linha.estreia.join(", ");
			linha["html-género"] = linha.género ?? "";
			// Formatação de nomes
			for (const nomeColuna of ["nomes", "nomes-originais", "nomes-inglês"]) {
				linha["html-" + nomeColuna] = linha[nomeColuna].formatarNomesHtml(linha["html-ordem"]);
			}
			// Formatação de outros nomes
			for (const nomeColuna of nomesColunasOutrosNomes) {
				linha["html-" + nomeColuna] = linha[nomeColuna].formatarListaNomesHtml(nomeTabela);
			}
			// Formatação das colunas de referência
			for (const nomeColuna of nomesColunasListasReferências) {
				linha["html-" + nomeColuna] = linha[nomeColuna + "-nomes"].formatarListaNomesHtml(nomeTabela);
			}
		}
	}
}

function preencherTabelaHtml() {
	for (const nomeTabela in tabelas) {
		const tabela = tabelas[nomeTabela];
		const tabelaHtml = document.getElementById(nomeTabela);
		const colunasTabelaHtml = [];
		const corpoTabelaHtml = tabelaHtml.tBodies[0];
		for (nomeColuna of nomesColunasHtml) {
			const colunaTabelaHtml = document.createElement("col");
			colunaTabelaHtml.dataset.tipo = nomeColuna;
			if (nomesColunasHtmlOcultas[nomeTabela].includes(nomeColuna)) {
				colunaTabelaHtml.dataset.oculto = "oculto";
			}
			colunasTabelaHtml.push(colunaTabelaHtml);
			tabelaHtml.prepend(...colunasTabelaHtml);
		}
		for (const id in tabela) {
			const linha = tabela[id];
			const linhaTabelaHtml = corpoTabelaHtml.insertRow();
			linhaTabelaHtml.id = id;
			if (linha["alter-ego"]) {
				linhaTabelaHtml.classList.add("alter-ego");
			}
			if (linha["identidade-secreta"]) {
				linhaTabelaHtml.classList.add("identidade-secreta");
			}
			if (linha["animal-estimação"]) {
				linhaTabelaHtml.classList.add("animal-estimação");
			}
			for (const nomeColuna of nomesColunasHtml) {
				if (nomeColuna === "número") {
					linhaTabelaHtml.insertCell();
				}
				else {
					linhaTabelaHtml.insertCell().innerHTML = linha["html-" + nomeColuna];
				}
			}
		}
		console.log(tabelaHtml);
	}
}

function listaNomes(ln, cols) {
	let res = [];
	for (const col of cols) {
		let item = ln[col];
		if (item) {
			item = item.toString();
		}
		res.push(item);
	}
	return res;
}

function preencherReferênciasCruzadas(ln, col, lnMatriz) {
	for (const i of ln[col]) {
		tabelas[lnMatriz[0]][lnMatriz[0] + "-" + i]?.[lnMatriz[1]].push(ln.código);
	}
}

function comparadorLinhas(a, b) {
	return a["nome-padrão"].prepararNomeOrdenação().localeCompare(b["nome-padrão"].prepararNomeOrdenação(), "pt-PT", {sensitivity: "base"});
}

function comparadorNomes(a, b) {
	return a[0].prepararNomeOrdenação().localeCompare(b[0].prepararNomeOrdenação(), "pt-PT", {sensitivity: "base"});

}

