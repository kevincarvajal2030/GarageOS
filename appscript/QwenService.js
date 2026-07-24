/**
 * ============================================================
 * GARAGE OS
 * QwenService.gs
 *
 * Central service responsible for communicating
 * with the Qwen (Alibaba Cloud DashScope) API.
 * Supports text generation and image generation.
 * ============================================================
 */

const QwenService = (() => {

  const BASE_URL = "https://dashscope-intl.aliyuncs.com/api/v1";

  /**
   * Returns the Qwen API Key stored in Script Properties.
   */
  function getApiKey() {

    const apiKey = PropertiesService
      .getScriptProperties()
      .getProperty("QWEN_API_KEY");

    if (!apiKey) {

      throw new Error(
        "Missing Script Property: QWEN_API_KEY"
      );

    }

    return apiKey;

  }

  /**
   * Sends an HTTP request to Qwen/DashScope API.
   */
  function request(endpoint, payload) {

    const response = UrlFetchApp.fetch(

      BASE_URL + endpoint,

      {

        method: "post",

        contentType: "application/json",

        headers: {

          Authorization: `Bearer ${getApiKey()}`

        },

        payload: JSON.stringify(payload),

        muteHttpExceptions: true

      }

    );

    return {

      status: response.getResponseCode(),

      body: JSON.parse(response.getContentText())

    };

  }

  /**
   * Parses a text response from Qwen.
   */
  function parseTextResponse(response) {

    if (response.status !== 200) {

      throw new Error(

        response.body.message ||
        response.body.error?.message ||
        "Unknown Qwen API error."

      );

    }

    if (!response.body.output || !response.body.output.text) {

      throw new Error(
        "Qwen returned no text."
      );

    }

    return response.body.output.text;

  }

  /**
   * Generates text using Qwen model.
   */
  function generateText(prompt) {

    const response = request(

      "/services/aigc/text-generation/generation",

      {

        model: QWEN.CHAT_MODEL,

        input: {
          messages: [
            {
              role: "user",
              content: prompt
            }
          ]
        },

        parameters: {
          result_format: "text"
        }

      }

    );

    return parseTextResponse(response);

  }

  /**
   * Generates an image using Qwen/DashScope API.
   * Uses the wanx-v1 model for image generation (available in free tier).
   */
  function generateImage(prompt, options = {}) {

    // Formato correcto: "1024*1024" no "1024x1024"
    const size = options.size || "1024*1024";
    const n = options.n || 1;

    // Parse size into width and height (soporta ambos formatos: "1024*1024" o "1024x1024")
    let width, height;
    if (size.includes("x")) {
      [width, height] = size.split("x").map(Number);
    } else {
      [width, height] = size.split("*").map(Number);
    }

    // Formato requerido por la API: "ancho*alto"
    const formattedSize = `${width}*${height}`;

    Logger.log(`Generando imagen con modelo ${QWEN.IMAGE_MODEL}, tamaño: ${formattedSize}`);

    const response = request(

      "/services/aigc/image-generation/generation",

      {

        model: QWEN.IMAGE_MODEL,

        input: {
          prompt: prompt
        },

        parameters: {
          size: formattedSize,
          n: n,
          style: "<auto>"
        }

      }

    );

    Logger.log(`Respuesta inicial - Status: ${response.status}`);

    if (response.status !== 200) {

      const errorMsg = response.body.message ||
                       response.body.error?.message ||
                       JSON.stringify(response.body);

      Logger.log(`Error en generación de imagen: ${errorMsg}`);

      throw new Error(errorMsg);

    }

    // Qwen returns task_id for async image generation
    // Need to poll for results
    if (response.body.output && response.body.output.task_id) {

      Logger.log(`Task ID recibido: ${response.body.output.task_id}. Esperando resultados...`);
      return waitForImageTask(response.body.output.task_id);

    }

    return response.body;

  }

  /**
   * Polls for image generation task completion.
   */
  function waitForImageTask(taskId, maxAttempts = 30, delayMs = 2000) {

    for (let i = 0; i < maxAttempts; i++) {

      const response = UrlFetchApp.fetch(

        BASE_URL + `/tasks/${taskId}`,

        {

          method: "get",

          contentType: "application/json",

          headers: {

            Authorization: `Bearer ${getApiKey()}`

          },

          muteHttpExceptions: true

        }

      );

      const result = JSON.parse(response.getContentText());

      if (result.output && result.output.task_status === "SUCCEEDED") {

        // Return the generated image URLs
        return result.output.results || [];

      } else if (result.output && result.output.task_status === "FAILED") {

        throw new Error(
          result.output.message || "Image generation task failed."
        );

      }

      // Wait before next poll
      Utilities.sleep(delayMs);

    }

    throw new Error("Image generation task timed out.");

  }

  /**
   * Simple API connectivity test.
   */
  function healthCheck() {

    try {

      const result = generateText(
        "Reply only with: OK"
      );

      return {
        status: "healthy",
        message: result
      };

    } catch (e) {

      return {
        status: "unhealthy",
        message: e.message
      };

    }

  }

  /**
   * Test image generation functionality.
   */
  function testImageGeneration() {

    try {

      const prompt = "A red sports car in a modern garage, photorealistic, high quality";

      Logger.log("Generating test image with prompt: " + prompt);

      const images = generateImage(prompt);

      Logger.log("Image generation successful!");
      Logger.log("Generated images: " + JSON.stringify(images));

      return {
        success: true,
        images: images
      };

    } catch (e) {

      Logger.log("Image generation failed: " + e.message);

      return {
        success: false,
        error: e.message
      };

    }

  }

  return Object.freeze({

    generateText,
    generateImage,
    healthCheck,
    testImageGeneration

  });

})();