/**
 * Integração com Gemini para categorização inteligente
 */
function categorizeTransactionWithAI(rawText) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error("API Key do Gemini não configurada.");
  
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
  
  const prompt = `Analise o seguinte texto bruto de extrato bancário: "${rawText}"
Retorne ESTRITAMENTE um JSON válido com duas chaves:
- "descricao": O nome limpo e legível do estabelecimento ou transação (ex: '99 Pop', 'Uber', 'Mercado Livre').
- "categoria": Escolha EXATAMENTE UMA desta lista: Alimentação, Transporte, Moradia, Lazer, Saúde, Renda, Transferência. Se não tiver certeza absoluta, retorne 'DESCONHECIDO'.
Não inclua crases, blocos de código markdown ou texto extra, apenas o JSON.`;

  const payload = {
    "contents": [{
      "parts": [{"text": prompt}]
    }]
  };
  
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    if (json.candidates && json.candidates.length > 0) {
      const textResponse = json.candidates[0].content.parts[0].text.trim();
      const cleanedJson = textResponse.replace(/^```json\n?|```$/g, '').trim();
      return JSON.parse(cleanedJson);
    }
    return { descricao: rawText, categoria: 'DESCONHECIDO' };
  } catch (e) {
    return { descricao: rawText, categoria: 'DESCONHECIDO' };
  }
}
