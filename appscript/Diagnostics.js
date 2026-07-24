
// Diagnostics.js
// Archivo para pruebas y diagnóstico de los servicios de IA.
//
// SERVICIOS ACTIVOS:
// - Pollinations.ai (Imágenes Gratuitas - SIN API KEY)

/**
 * Ejecuta todas las pruebas de los servicios activos actualmente.
 */
function runAllActiveTests() {
  Logger.log("=== INICIANDO SUITE DE PRUEBAS ACTIVA ===");

  try {
    testPollinationsGeneration();
  } catch (e) {
    Logger.log("Error en prueba Pollinations: " + e.toString());
  }

  Logger.log("=== SUITE DE PRUEBAS FINALIZADA ===");
}

/**
 * =============================================================================
 * PRUEBAS PARA POLLINATIONS.AI (IMÁGENES GRATUITAS)
 * =============================================================================
 */

/**
 * Prueba básica de generación de imagen con Pollinations.
 * No requiere API Key.
 */
function testPollinationsGeneration() {
  Logger.log("--- Iniciando prueba de generación de imagen (Pollinations) ---");

  var prompt = "A futuristic cyberpunk city with neon lights, realistic, 8k";
  Logger.log("Generando imagen con prompt: '" + prompt + "'");

  try {
    var result = PollinationsService.generateImage(prompt, {
      width: 1024,
      height: 1024,
      model: 'flux' // Modelo recomendado para calidad/velocidad
    });

    if (result.success) {
      Logger.log("¡ÉXITO! Imagen generada correctamente.");
      Logger.log("URL de la imagen: " + result.url);
      Logger.log("Modelo utilizado: " + (result.model || 'default'));
      Logger.log("Haz clic aquí para ver la imagen: " + result.url);
    } else {
      Logger.log("Fallo en la generación: " + (result.error || "Error desconocido"));
    }
  } catch (e) {
    Logger.log("Excepción durante la generación: " + e.toString());
  }
}

/**
 * Prueba de generación y guardado en Google Drive.
 * Requiere permisos de escritura en Drive.
 */
function testPollinationsToDrive() {
  Logger.log("--- Iniciando prueba de guardado en Drive (Pollinations) ---");

  var prompt = "A cute robot holding a flower in a garden, Pixar style";
  Logger.log("Generando y guardando imagen: '" + prompt + "'");

  try {
    // Intenta guardar en la carpeta raíz o una específica si se proporciona ID
    var folderId = null; // Deja null para usar la raíz, o pon un ID de carpeta

    var result = PollinationsService.generateImageToDrive(prompt, {
      width: 1024,
      height: 1024,
      model: 'flux-realism'
    }, folderId);

    if (result.success) {
      Logger.log("¡ÉXITO! Imagen guardada en Drive.");
      Logger.log("Nombre del archivo: " + result.fileName);
      Logger.log("ID del archivo: " + result.fileId);
      Logger.log("URL de visualización: " + result.viewUrl);
    } else {
      Logger.log("Fallo al guardar en Drive: " + (result.error || "Error desconocido"));
    }
  } catch (e) {
    Logger.log("Excepción al guardar en Drive: " + e.toString());
  }
}

/**
 * Prueba iterativa de diferentes modelos disponibles en Pollinations.
 */
function testPollinationsModels() {
  Logger.log("--- Probando diferentes modelos de Pollinations ---");

  var models = ['flux', 'flux-realism', 'stable-diffusion-xl', 'turbo'];
  var prompt = "A golden trophy on a pedestal";

  models.forEach(function(model) {
    Logger.log("Probando modelo: " + model);
    try {
      var result = PollinationsService.generateImage(prompt, {
        width: 512, // Tamaño reducido para prueba rápida
        height: 512,
        model: model
      });

      if (result.success) {
        Logger.log("  -> " + model + ": OK (URL: " + result.url + ")");
      } else {
        Logger.log("  -> " + model + ": FALLÓ (" + result.error + ")");
      }
    } catch (e) {
      Logger.log("  -> " + model + ": ERROR (" + e.toString() + ")");
    }
  });

  Logger.log("Prueba de modelos finalizada.");
}

/**
 * Prueba de generación de imagen de vehículo con Pollinations.
 */
function testVehicleImageWithPollinations() {
  Logger.log("--- Iniciando prueba de imagen de vehículo (Pollinations) ---");

  var vehicle = {
    make: "Toyota",
    model: "Supra",
    year: 2024,
    color: "Red"
  };

  var prompt = PromptService.buildVehiclePrompt(vehicle);
  Logger.log("Prompt generado: " + prompt);

  try {
    var result = PollinationsService.generateImage(prompt, {
      width: 1024,
      height: 1024,
      model: 'flux-realism'
    });

    if (result.success) {
      Logger.log("¡ÉXITO! Imagen de vehículo generada.");
      Logger.log("URL: " + result.url);
    } else {
      Logger.log("Fallo: " + (result.error || "Error desconocido"));
    }
  } catch (e) {
    Logger.log("Excepción: " + e.toString());
  }
}