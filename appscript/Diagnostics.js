//Diagnostics.gs

function testGenerateText() {

  const text = OpenAIService.generateText(

    "Respond only with: GarageOS connection successful."

  );

  Logger.log(text);

}


function testHealthCheck() {

    Logger.log(

        OpenAIService.healthCheck()

    );

}



function testGenerateImage() {

  const vehicle = {

    make: "Toyota",
    model: "Supra",
    year: 2024,
    color: "Red"

  };

  const prompt =
    PromptService.buildVehiclePrompt(vehicle);

  const response =
    OpenAIService.generateImage(prompt);

  Logger.log(

    JSON.stringify(response, null, 2)

  );

}



