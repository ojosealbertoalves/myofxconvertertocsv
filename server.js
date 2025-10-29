const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// Função para formatar a data do OFX (formato: 20251027185400[-3:EST]) para DD/MM/YYYY
function formatarData(dataOfx) {
  if (!dataOfx) return '';
  
  const dataStr = dataOfx.toString().trim();
  const ano = dataStr.substring(0, 4);
  const mes = dataStr.substring(4, 6);
  const dia = dataStr.substring(6, 8);
  
  return `${dia}/${mes}/${ano}`;
}

// Função para fazer parse manual do OFX
function parseOfx(conteudo) {
  const transacoes = [];
  
  // Dividir o conteúdo em blocos de transação
  const blocos = conteudo.split('<STMTTRN>');
  
  // Remover o primeiro bloco (cabeçalho)
  blocos.shift();
  
  blocos.forEach(bloco => {
    const transacao = {};
    
    // Extrair TRNTYPE
    const tipoMatch = bloco.match(/<TRNTYPE>([^\n<]+)/);
    if (tipoMatch) transacao.TRNTYPE = tipoMatch[1].trim();
    
    // Extrair DTPOSTED
    const dataMatch = bloco.match(/<DTPOSTED>([^\n<]+)/);
    if (dataMatch) transacao.DTPOSTED = dataMatch[1].trim();
    
    // Extrair TRNAMT
    const valorMatch = bloco.match(/<TRNAMT>([^\n<]+)/);
    if (valorMatch) transacao.TRNAMT = valorMatch[1].trim();
    
    // Extrair FITID
    const idMatch = bloco.match(/<FITID>([^\n<]+)/);
    if (idMatch) transacao.FITID = idMatch[1].trim();
    
    // Extrair MEMO
    const memoMatch = bloco.match(/<MEMO>([^\n<]+)/);
    if (memoMatch) transacao.MEMO = memoMatch[1].trim();
    
    // Adicionar à lista se tiver os campos essenciais
    if (transacao.DTPOSTED && transacao.TRNAMT) {
      transacoes.push(transacao);
    }
  });
  
  return transacoes;
}

// Função para converter OFX para array de objetos
function converterOfxParaDados(arquivoOfx) {
  const dadosOfx = fs.readFileSync(arquivoOfx, 'utf8');
  const transacoes = parseOfx(dadosOfx);
  
  // Formatar os dados
  const dadosFormatados = transacoes.map((t) => {
    return {
      'tipo': t.TRNTYPE, // CREDIT ou DEBIT
      'data': formatarData(t.DTPOSTED),
      'valor': parseFloat(t.TRNAMT),
      'descric': t.MEMO || '',
      'id': t.FITID || '',
      'checks': ''
    };
  });
  
  return dadosFormatados;
}

// Função para salvar como XLSX
function salvarComoXLSX(dados, arquivoSaida) {
  const worksheet = XLSX.utils.json_to_sheet(dados);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Transações');
  
  // Ajustar largura das colunas
  const colunas = [
    { wch: 10 }, // tipo
    { wch: 12 }, // data
    { wch: 15 }, // valor
    { wch: 60 }, // descric
    { wch: 80 }, // id
    { wch: 10 }  // checks
  ];
  worksheet['!cols'] = colunas;
  
  XLSX.writeFile(workbook, arquivoSaida);
}

// Função para salvar como CSV
function salvarComoCSV(dados, arquivoSaida) {
  // Criar cabeçalho
  const colunas = Object.keys(dados[0]);
  let csv = colunas.join(';') + '\n';
  
  // Adicionar linhas
  dados.forEach(linha => {
    const valores = colunas.map(coluna => {
      const valor = linha[coluna];
      // Se contiver ponto e vírgula ou quebra de linha, colocar entre aspas
      if (typeof valor === 'string' && (valor.includes(';') || valor.includes('\n'))) {
        return `"${valor}"`;
      }
      return valor;
    });
    csv += valores.join(';') + '\n';
  });
  
  fs.writeFileSync(arquivoSaida, csv, 'utf8');
}

// Função para processar todos os arquivos OFX de uma pasta
function processarTodosOfx(pastaOfx, pastaDestino) {
  console.log('='.repeat(60));
  console.log('🚀 CONVERSOR DE ARQUIVOS OFX PARA XLSX/CSV');
  console.log('='.repeat(60));
  console.log();
  
  // Criar pasta de destino se não existir
  if (!fs.existsSync(pastaDestino)) {
    fs.mkdirSync(pastaDestino, { recursive: true });
    console.log(`📁 Pasta de saída criada: ${pastaDestino}\n`);
  }
  
  // Verificar se a pasta de origem existe
  if (!fs.existsSync(pastaOfx)) {
    console.error(`❌ Erro: A pasta '${pastaOfx}' não existe!`);
    console.log('\n💡 Dica: Crie a pasta e coloque os arquivos .ofx dentro dela.');
    return;
  }
  
  // Ler todos os arquivos da pasta
  const arquivos = fs.readdirSync(pastaOfx);
  
  // Filtrar apenas arquivos .ofx
  const arquivosOfx = arquivos.filter(arquivo => 
    arquivo.toLowerCase().endsWith('.ofx')
  );
  
  if (arquivosOfx.length === 0) {
    console.log(`⚠️  Nenhum arquivo .ofx encontrado na pasta '${pastaOfx}'`);
    console.log('\n💡 Dica: Adicione arquivos .ofx na pasta e execute novamente.');
    return;
  }
  
  console.log(`📂 Encontrados ${arquivosOfx.length} arquivo(s) OFX para processar:\n`);
  
  let totalTransacoes = 0;
  let arquivosProcessados = 0;
  
  // Processar cada arquivo
  arquivosOfx.forEach((nomeArquivo, index) => {
    try {
      console.log(`[${index + 1}/${arquivosOfx.length}] Processando: ${nomeArquivo}`);
      
      const caminhoCompleto = path.join(pastaOfx, nomeArquivo);
      const dados = converterOfxParaDados(caminhoCompleto);
      
      // Calcular totais
      const totalCredito = dados
        .filter(t => t.tipo === 'CREDIT')
        .reduce((sum, t) => sum + t.valor, 0);
      
      const totalDebito = dados
        .filter(t => t.tipo === 'DEBIT')
        .reduce((sum, t) => sum + Math.abs(t.valor), 0);
      
      console.log(`   ✓ ${dados.length} transações encontradas`);
      console.log(`   ✓ Créditos: R$ ${totalCredito.toFixed(2)}`);
      console.log(`   ✓ Débitos: R$ ${totalDebito.toFixed(2)}`);
      console.log(`   ✓ Saldo: R$ ${(totalCredito - totalDebito).toFixed(2)}`);
      
      // Nome do arquivo de saída (sem extensão .ofx)
      const nomeBase = path.basename(nomeArquivo, '.ofx');
      
      // Salvar XLSX
      const caminhoXlsx = path.join(pastaDestino, `${nomeBase}.xlsx`);
      salvarComoXLSX(dados, caminhoXlsx);
      console.log(`   ✓ Salvo: ${nomeBase}.xlsx`);
      
      // Salvar CSV
      const caminhoCsv = path.join(pastaDestino, `${nomeBase}.csv`);
      salvarComoCSV(dados, caminhoCsv);
      console.log(`   ✓ Salvo: ${nomeBase}.csv`);
      
      console.log();
      
      totalTransacoes += dados.length;
      arquivosProcessados++;
      
    } catch (erro) {
      console.error(`   ❌ Erro ao processar ${nomeArquivo}:`, erro.message);
      console.log();
    }
  });
  
  console.log('='.repeat(60));
  console.log('📊 RESUMO FINAL');
  console.log('='.repeat(60));
  console.log(`✅ Arquivos processados: ${arquivosProcessados}/${arquivosOfx.length}`);
  console.log(`📈 Total de transações convertidas: ${totalTransacoes}`);
  console.log(`📁 Arquivos salvos em: ${pastaDestino}`);
  console.log('='.repeat(60));
}

// Função principal
function main() {
  // Configurações
  const pastaOfx = './arquivos-ofx';      // Pasta onde estão os arquivos .ofx
  const pastaDestino = './convertidos';   // Pasta onde serão salvos os arquivos convertidos
  
  processarTodosOfx(pastaOfx, pastaDestino);
}

// Executar
main();