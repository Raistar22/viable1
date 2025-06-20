function testGeminiApiKeyRetrieval() {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (apiKey) {
      Logger.log('GEMINI_API_KEY was retrieved successfully: ' + apiKey.substring(0, 5) + '...'); // Log first 5 chars for safety
    } else {
      Logger.log('GEMINI_API_KEY is NULL or undefined.');
    }
  } catch (e) {
    Logger.log('An error occurred during key retrieval: ' + e.message);
  }
}