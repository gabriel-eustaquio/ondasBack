import express from 'express';
import { fileURLToPath } from 'url';
import { dirname } from 'path';
import XlsxPopulate from 'xlsx-populate';
import cors from 'cors';

const app = express();
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
app.use(
  cors({
    origin: 'https://gabriel-eustaquio.github.io/ondas/',
  }),
);

app.use(express.json());

app.post('/adicionar', async (req, res) => {
  try {
    const workbook = await XlsxPopulate.fromFileAsync('./cadastro.xlsx');
    const sheet = workbook.sheet(1);
    console.log(req.body);

    const {
      name,
      email,
      telephone,
      childs,
      enrollmentOrReplacement,
      preferTime,
      enrollmentBy,
      voluntary,
    } = req.body;

    if (
      !name ||
      !email ||
      !childs ||
      !enrollmentOrReplacement ||
      !preferTime ||
      !enrollmentBy ||
      !voluntary
    ) {
      throw new Error('Dados inválidos');
    }

    // Encontrar a última linha com dados na coluna A (assumindo que A é a coluna-chave)
    let lastRow = 1;
    while (sheet.cell(lastRow, 1).value()) {
      lastRow++;
    }

    console.log('Quantidade de registro:', lastRow);

    // Adicionar os dados à planilha na próxima linha após a última linha com dados
    sheet.cell(lastRow, 1).value(name);
    sheet.cell(lastRow, 3).value(email);
    sheet.cell(lastRow, 5).value(telephone);
    sheet.cell(lastRow, 7).value(childs);
    sheet.cell(lastRow, 9).value(enrollmentOrReplacement);
    sheet.cell(lastRow, 11).value(preferTime);
    sheet.cell(lastRow, 13).value(enrollmentBy);
    sheet.cell(lastRow, 15).value(voluntary);

    // Manter estilos e filtros
    const range = sheet.range(`A1:C${lastRow}`);
    const definedName = workbook.definedName('_xlnm._FilterDatabase');
    if (definedName) {
      definedName.value(range);
    }

    await workbook.toFileAsync('./cadastro.xlsx');

    res.send('Dados adicionados com sucesso.');
  } catch (error) {
    console.error('Erro ao adicionar dados ao arquivo XLSX:', error);
    res.status(500).send('Erro ao adicionar dados ao arquivo XLSX.');
  }
});

app.post('/adicionarCheckin', async (req, res) => {
  try {
    const workbook = await XlsxPopulate.fromFileAsync('./checkin.xlsx');
    const sheet = workbook.sheet(0);
    console.log(req.body);

    const { name, amountChilds, nameChilds, teacher, currentDateTime } =
      req.body;

    if (!name || !amountChilds || !nameChilds || !teacher || !currentDateTime) {
      throw new Error('Dados inválidos');
    }

    // Encontrar a última linha com dados na coluna A (assumindo que A é a coluna-chave)
    let lastRow = 1;
    while (sheet.cell(lastRow, 1).value()) {
      lastRow++;
    }

    console.log('Quantidade de registro:', lastRow);

    // Adicionar os dados à planilha na próxima linha após a última linha com dados
    sheet.cell(lastRow, 1).value(name);
    sheet.cell(lastRow, 2).value(currentDateTime);
    sheet.cell(lastRow, 3).value(amountChilds);
    sheet.cell(lastRow, 4).value(nameChilds);
    sheet.cell(lastRow, 5).value(teacher);

    // Manter estilos e filtros
    const range = sheet.range(`A1:C${lastRow}`);
    const definedName = workbook.definedName('_xlnm._FilterDatabase');
    if (definedName) {
      definedName.value(range);
    }

    await workbook.toFileAsync('./checkin.xlsx');

    res.send('Dados adicionados com sucesso.');
  } catch (error) {
    console.error('Erro ao adicionar dados ao arquivo XLSX:', error);
    res.status(500).send('Erro ao adicionar dados ao arquivo XLSX.');
  }
});

app.get('/download', (req, res) => {
  try {
    res.download('./cadastro.xlsx', 'cadastro.xlsx');
  } catch (error) {
    console.error('Erro ao fazer download do arquivo XLSX:', error);
    res.status(500).send('Erro ao fazer download do arquivo XLSX.');
  }
});

app.get('/downloadCheckin', (req, res) => {
  try {
    res.download('./checkin.xlsx', 'checkin.xlsx');
  } catch (error) {
    console.error('Erro ao fazer download do arquivo XLSX:', error);
    res.status(500).send('Erro ao fazer download do arquivo XLSX.');
  }
});

app.listen(5000, () => {
  console.log('Servidor iniciado.');
});
