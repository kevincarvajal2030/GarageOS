/**
 * ============================================================
 * GARAGE OS
 * PromptService.gs
 *
 * Central service responsible for building AI prompts.
 *
 * This service NEVER communicates with Gemini.
 * It only builds prompt strings.
 * ============================================================
 */

const PromptService = (() => {

  /**
   * Builds the vehicle description section.
   */
  function buildVehicleDescription(vehicle) {

    const details = [

      `Make: ${vehicle.make}`,
      `Model: ${vehicle.model}`,
      `Year: ${vehicle.year}`,
      `Color: ${vehicle.color}`

    ];

    return [

      "Vehicle:",
      "",
      ...details.map(detail => `- ${detail}`)

    ].join("\n");

  }


  /**
   * Builds the image generation instructions.
   */
  function buildImageInstructions() {

    return [

      "Requirements:",
      "",
      "- Professional automotive studio photograph.",
      "- Front three-quarter view.",
      "- Entire vehicle visible.",
      "- White seamless background.",
      "- Centered composition.",
      "- Photorealistic.",
      "- Ultra realistic materials.",
      "- High detail.",
      "- Sharp focus.",
      "- Modern studio lighting.",
      "- Realistic reflections.",
      "- Factory stock appearance.",
      "- No people.",
      "- No buildings.",
      "- No roads.",
      "- No logos.",
      "- No license plate text.",
      "- No watermark.",
      "- No text."

    ].join("\n");

  }


  /**
   * Builds the complete prompt for generating
   * a vehicle image.
   */
  function buildVehiclePrompt(vehicle) {

    return [

      buildVehicleDescription(vehicle),
      "",
      buildImageInstructions()

    ].join("\n");

  }


  return Object.freeze({

    buildVehiclePrompt

  });

})();