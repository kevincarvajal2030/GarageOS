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



