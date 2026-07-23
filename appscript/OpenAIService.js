/**
 * ============================================================
 * GARAGE OS
 * OpenAIService.gs
 *
 * Central service responsible for communicating
 * with the OpenAI API.
 * ============================================================
 */

const OpenAIService = (() => {

  const BASE_URL = "https://api.openai.com/v1";

  /**
   * Returns the OpenAI API Key stored in Script Properties.
   */
  function getApiKey() {

    const apiKey = PropertiesService
      .getScriptProperties()
      .getProperty("OPENAI_API_KEY");

    if (!apiKey) {

      throw new Error(
        "Missing Script Property: OPENAI_API_KEY"
      );

    }

    return apiKey;

  }

  /**
   * Sends an HTTP request to OpenAI.
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
   * Parses a text response.
   */
  function parseTextResponse(response) {

    if (response.status !== 200) {

      throw new Error(

        response.body.error?.message ||
        "Unknown OpenAI API error."

      );

    }

    if (!response.body.output_text) {

      throw new Error(
        "OpenAI returned no text."
      );

    }

    return response.body.output_text;

  }

  /**
   * Generates text.
   */
  function generateText(prompt) {

    const response = request(

      "/responses",

      {

        model: OPENAI.CHAT_MODEL,

        input: prompt

      }

    );

    return parseTextResponse(response);

  }

  /**
   * Simple API connectivity test.
   */
  function healthCheck() {

    return generateText(

      "Reply only with: OK"

    );

  }

  /**
   * Generates an image.
   *
   * Image parsing will be implemented later.
   */
  function generateImage(prompt) {

    return request(

      "/images/generations",

      {

        model: OPENAI.IMAGE_MODEL,

        prompt: prompt,

        size: "1024x1024"

      }

    );

  }

  return Object.freeze({

    generateText,
    generateImage,
    healthCheck

  });

})();