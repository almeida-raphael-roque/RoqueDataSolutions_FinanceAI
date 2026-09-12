
function categorizeBatchWithAI(descriptions, availableCategories) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error("API Key do Gemini não configurada.");
  
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
  
  const prompt = `Analise os seguintes textos brutos de extrato bancário.
Para cada um, identifique o estabelecimento, sugira uma categoria e forneça um nível de confiança (0.0 a 1.0).
Categorias disponíveis: ${availableCategories.join(', ')}.

Textos:
${descriptions.map((d, i) => `[ID:${i}] ${d}`).join('\n')}

Retorne ESTRITAMENTE um JSON válido com a seguinte estrutura:
{
  "resultados": [
    {
      "id": 0,
      "estabelecimento": "Nome limpo",
      "categoria": "Categoria sugerida",
      "confianca": 0.95
    }
  ]
}
Não inclua texto adicional.`;

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
    return { resultados: [] };
  } catch (e) {
    return { resultados: [] };
  }
}
