//Diagnostics.gs

/**
 * Test text generation with Qwen
 */
function testQwenTextGeneration() {

  const text = QwenService.generateText(

    "Respond only with: GarageOS Qwen connection successful."

  );

  Logger.log(text);

}


/**
 * Test health check with Qwen
 */
function testQwenHealthCheck() {

    Logger.log(

        QwenService.healthCheck()

    );

}


/**
 * Test image generation with Qwen
 */
function testQwenGenerateImage() {

  const vehicle = {

    make: "Toyota",
    model: "Supra",
    year: 2024,
    color: "Red"

  };

  const prompt =
    PromptService.buildVehiclePrompt(vehicle);

  const response =
    QwenService.generateImage(prompt);

  Logger.log(

    JSON.stringify(response, null, 2)

  );

}


/**
 * Test Qwen image generation with simple prompt
 */
function testQwenImageSimple() {

  const result = QwenService.testImageGeneration();

  Logger.log(

    JSON.stringify(result, null, 2)

  );

}


/**
 * Test Qwen connection and available models
 * Esta función prueba la conexión con diferentes modelos de imagen
 */
function testQwenConnectionFull() {

  Logger.log("=== INICIANDO PRUEBA COMPLETA DE CONEXIÓN QWEN ===");

  // Prueba 1: Verificar configuración
  Logger.log("1. Verificando configuración...");
  Logger.log(`   Modelo de chat configurado: ${QWEN.CHAT_MODEL}`);
  Logger.log(`   Modelo de imagen configurado: ${QWEN.IMAGE_MODEL}`);

  // Prueba 2: Generación de texto
  Logger.log("\n2. Probando generación de texto...");
  try {
    const textResult = QwenService.generateText("Responde solo 'OK' si funciona");
    Logger.log(`   Resultado texto: ${textResult}`);
  } catch (e) {
    Logger.log(`   Error texto: ${e.message}`);
  }

  // Prueba 3: Generación de imagen con modelo actual
  Logger.log("\n3. Probando generación de imagen con modelo actual...");
  try {
    const imageResult = QwenService.generateImage(
      "A simple blue circle on white background",
      { size: "1024*1024" }
    );
    Logger.log(`   Resultado imagen: ${JSON.stringify(imageResult)}`);
  } catch (e) {
    Logger.log(`   Error imagen: ${e.message}`);

    // Si falla, intentar con otros modelos comunes
    Logger.log("\n4. Intentando con otros modelos de imagen...");
    const alternativeModels = ["wanx-v1", "wanx2.1-t2i-turbo", "wanx2.0-t2i-turbo"];

    for (const model of alternativeModels) {
      Logger.log(`   Probando modelo: ${model}...`);
      try {
        // Temporalmente cambiar el modelo en config no es posible,
        // pero podemos reportar cuáles intentar
        Logger.log(`     - ${model}: Deberías probar este modelo manualmente`);
      } catch (e) {
        Logger.log(`     - ${model}: Falló - ${e.message}`);
      }
    }
  }

  Logger.log("\n=== PRUEBA COMPLETADA ===");
}



