const XLSX = require("xlsx");
const axios = require("axios");
const fs = require("fs");

const brazilStates = {
  AC: "Acre",
  AL: "Alagoas",
  AP: "Amapá",
  AM: "Amazonas",
  BA: "Bahia",
  CE: "Ceará",
  DF: "Distrito Federal",
  ES: "Espírito Santo",
  GO: "Goiás",
  MA: "Maranhão",
  MT: "Mato Grosso",
  MS: "Mato Grosso do Sul",
  MG: "Minas Gerais",
  PA: "Pará",
  PB: "Paraíba",
  PR: "Paraná",
  PE: "Pernambuco",
  PI: "Piauí",
  RJ: "Rio de Janeiro",
  RN: "Rio Grande do Norte",
  RS: "Rio Grande do Sul",
  RO: "Rondônia",
  RR: "Roraima",
  SC: "Santa Catarina",
  SP: "São Paulo",
  SE: "Sergipe",
  TO: "Tocantins",
};

const preposicoes = [
  "a",
  "ao",
  "aos",
  "ante",
  "ate",
  "até",
  "apos",
  "após",
  "com",
  "contra",
  "em",
  "entre",
  "para",
  "por",
  "perante",
  "sem",
  "sob",
  "sobre",
  "tras",
  "trás",
];
const regexPreposicoes = new RegExp(
  `(^|\\s)(?:${preposicoes.join("|")})(\\s|$)`
);

let url =
  "https://docs.google.com/spreadsheets/d/1Ln_v12Zjf-w0l6h9_Nh9jE27k4JimqXWzh-y7AVwqcI/edit?usp=drivesdk";

function correctText(text) {
  return decodeURIComponent(escape(text));
}

async function buscaEnderecoPorCEP(cep) {
  try {
    const response = await axios.get(`https://viacep.com.br/ws/${cep}/json/`);
    const endereco = response.data;
    return {
      rua: endereco.logradouro || "Rua não encontrada",
      bairro: endereco.bairro || "Bairro não encontrado",
      cidade: endereco.localidade || "Cidade não encontrada",
      estado: endereco.uf || "Estado não encontrado",
    };
  } catch (error) {
    console.error(error);
    return {
      rua: "Rua não encontrada",
      bairro: "Bairro não encontrado",
      cidade: "Cidade não encontrada",
      estado: "Estado não encontrado",
    };
  }
}

axios({
  url,
  method: "GET",
  responseType: "stream",
}).then((response) => {
  let path = "./dataset.xlsx";
  let writer = fs.createWriteStream(path);
  response.data.pipe(writer);

  writer.on("finish", () => {
    let workbook = XLSX.readFile(path);
    let sheetName = workbook.SheetNames[0];
    let worksheet = workbook.Sheets[sheetName];
    let data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    let processedData = data
      .slice(2)
      .map(async (row) => {
        const dataFromTable = row[1];
        if (dataFromTable) {
          let endereco = correctText(dataFromTable);

          let ruaNumeroMatch = endereco.match(/^(.+?),\s*([\d-]+)?/);
          let rua = ruaNumeroMatch
            ? ruaNumeroMatch[1].trim()
            : "Rua não encontrada";
          let numero =
            ruaNumeroMatch && ruaNumeroMatch[2]
              ? ruaNumeroMatch[2].trim()
              : "Número não encontrado";

          let bairroMatch = endereco
            .substr(rua.length + numero.length)
            .match(/-\s*(.+?),\s*([^,]+)\s*-\s*([A-Z]{2}),/);
          let bairro = bairroMatch
            ? bairroMatch[1].trim()
            : "Bairro não encontrado";

          let cidadeEstadoMatch = endereco.match(
            /,\s*([^,]+)\s*-\s*([A-Z]{2}),/
          );
          let cidade = cidadeEstadoMatch
            ? cidadeEstadoMatch[1].trim()
            : "Cidade não encontrada";
          let estado = cidadeEstadoMatch
            ? brazilStates[cidadeEstadoMatch[2].trim().toUpperCase()] ||
              "Estado não encontrado"
            : "Estado não encontrado";

          let cepMatch = endereco.match(/(\d{5}-\d{3})/);
          let cep = cepMatch ? cepMatch[1] : "CEP não encontrado";

          let paisMatch = endereco.match(/,\s*([^,]+)$/);
          let pais = paisMatch ? paisMatch[1].trim() : "País não encontrado";

          // Caso não encontre um número na rua, tente encontrar a rua de novo sem um número
          if (numero === "Número não encontrado") {
            let ruaSemNumeroMatch = endereco.match(/^(.+?)(?=\s*,)/);
            rua = ruaSemNumeroMatch
              ? ruaSemNumeroMatch[1].trim()
              : "Rua não encontrada";
          }

          // Se não encontrar um bairro, tente encontrar a cidade e o estado de novo sem um bairro
          if (bairro === "Bairro não encontrado") {
            let cidadeEstadoSemBairroMatch = endereco.match(
              /,\s*([^,]+)\s*-\s*([A-Z]{2})(?=\s*,\s*\d{5}-\d{3})/
            );
            cidade = cidadeEstadoSemBairroMatch
              ? cidadeEstadoSemBairroMatch[1].trim()
              : "Cidade não encontrada";
            estado = cidadeEstadoSemBairroMatch
              ? brazilStates[
                  cidadeEstadoSemBairroMatch[2].trim().toUpperCase()
                ] || "Estado não encontrado"
              : "Estado não encontrado";
          }

          // Se a rua e o número forem encontrados na variável 'bairro', defina 'bairro' como "Bairro não encontrado"
          if (bairro.includes(rua) && bairro.includes(numero)) {
            bairro = "Bairro não encontrado";
          }

          // Caso a rua seja "Unnamed Road", defina como "Rua não encontrada"
          if (rua === "Unnamed Road") {
            rua = "Rua não encontrada";
          }

          // Corte as informações extras baseado na primeira preposição que aparece na rua
          let extraMatch = rua.toLowerCase().match(regexPreposicoes);
          let extra = "Extra não encontrado";
          if (extraMatch) {
            const index =
              extraMatch.index + (extraMatch[0].startsWith(" ") ? 1 : 0);
            extra = rua.slice(index).trim();
            rua = rua.slice(0, index).trim();
          }

          if (
            rua === "Rua não encontrada" ||
            bairro === "Bairro não encontrado" ||
            cidade === "Cidade não encontrada" ||
            estado === "Estado não encontrado"
          ) {
            const enderecoPorCEP = await buscaEnderecoPorCEP(cep);
            rua = rua === "Rua não encontrada" ? enderecoPorCEP.rua : rua;
            bairro =
              bairro === "Bairro não encontrada"
                ? enderecoPorCEP.bairro
                : bairro;
            cidade =
              cidade === "Cidade não encontrada"
                ? enderecoPorCEP.cidade
                : cidade;
            estado =
              estado === "Estado não encontrado"
                ? enderecoPorCEP.estado
                : estado;
          }

          return {
            Rua: rua,
            Numero: numero,
            Bairro: bairro,
            CEP: cep,
            Cidade: cidade,
            Estado: estado,
            Pais: pais,
            Extra: extra,
          };
        } else {
          return null;
        }
      })
      .filter((item) => item !== null);

    Promise.all(processedData).then((resolvedData) => {
      fs.writeFileSync(
        "dataset_processado.json",
        JSON.stringify(resolvedData, null, 2)
      );
      console.log("Excel file successfully processed");
    });
  });

  writer.on("error", console.error);
});
