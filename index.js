const fs = require("fs");
const XLSX = require("xlsx");

// Caminho do arquivo de entrada e saída
const inputFile = "numeros.txt"; // Substitua pelo nome do seu arquivo
const outputFile = "numeros.xlsx";

// Função para converter a lista em Excel
function convertTxtToExcel() {
  try {
    // Ler o conteúdo do arquivo .txt
    const data = fs.readFileSync(inputFile, "utf-8");

    // Dividir os números pelo separador (vírgula e espaços)
    const numbers = data.split(",").map((num) => num.trim());

    // Criar um array de objetos para o Excel
    const rows = numbers.map((number) => ({ Número: number }));

    // Criar uma planilha a partir do array
    const worksheet = XLSX.utils.json_to_sheet(rows);

    // Criar uma pasta de trabalho e adicionar a planilha
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Números");

    // Salvar como arquivo Excel
    XLSX.writeFile(workbook, outputFile);

    console.log(`Arquivo Excel criado com sucesso: ${outputFile}`);
  } catch (error) {
    console.error("Erro ao processar o arquivo:", error.message);
  }
}

// Executar o script
convertTxtToExcel();