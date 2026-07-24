
/**
 * Servicio de Generación de Imágenes con Pollinations.ai
 * API GRATUITA - Sin requerir API Key, sin registro, sin límites estrictos
 *
 * Características:
 * - Totalmente gratuito
 * - No requiere autenticación
 * - Modelos basados en Stable Diffusion XL
 * - Respuesta directa en formato URL o base64
 *
 * Documentación: https://pollinations.ai/docs
 */

var PollinationsService = (function() {

  var BASE_URL = 'https://image.pollinations.ai/prompt';

  /**
   * Genera una imagen usando Pollinations.ai
   * @param {string} prompt - Descripción de la imagen a generar
   * @param {object} options - Opciones adicionales
   *   - width: número (default: 1024)
   *   - height: número (default: 1024)
   *   - seed: número (opcional, para reproducibilidad)
   *   - model: string (opcional, default: 'flux')
   *   - nologo: boolean (default: true, sin marca de agua)
   * @return {object} Resultado con URL de la imagen
   */
  function generateImage(prompt, options) {
    try {
      Logger.log('=== GENERANDO IMAGEN CON POLLINATIONS.AI ===');
      Logger.log('Prompt: ' + prompt);

      // Valores por defecto
      var width = options && options.width ? options.width : 1024;
      var height = options && options.height ? options.height : 1024;
      var seed = options && options.seed ? options.seed : Math.floor(Math.random() * 1000000);
      var model = options && options.model ? options.model : 'flux';
      var nologo = options && options.nologo !== undefined ? options.nologo : true;

      // Construir URL con parámetros
      var encodedPrompt = encodeURIComponent(prompt);
      var imageUrl = BASE_URL + '/' + encodedPrompt +
                     '?width=' + width +
                     '&height=' + height +
                     '&seed=' + seed +
                     '&model=' + model +
                     '&nologo=' + nologo;

      Logger.log('URL generada: ' + imageUrl);
      Logger.log('Modelo: ' + model + ', Tamaño: ' + width + 'x' + height);

      // Verificar que la URL es accesible (HEAD request)
      var response = UrlFetchApp.fetch(imageUrl, {
        method: 'get',
        muteHttpExceptions: true
      });

      var responseCode = response.getResponseCode();
      Logger.log('Código de respuesta: ' + responseCode);

      if (responseCode === 200) {
        // Obtener la imagen como blob
        var imageBlob = response.getBlob();
        var contentType = imageBlob.getContentType();

        Logger.log('Imagen obtenida exitosamente');
        Logger.log('Tipo de contenido: ' + contentType);
        Logger.log('Tamaño del archivo: ' + Math.round(imageBlob.getBytes().length / 1024) + ' KB');

        return {
          success: true,
          url: imageUrl,
          blob: imageBlob,
          contentType: contentType,
          size: imageBlob.getBytes().length,
          prompt: prompt,
          model: model,
          dimensions: width + 'x' + height,
          seed: seed
        };
      } else {
        Logger.log('Error al obtener imagen: HTTP ' + responseCode);
        return {
          success: false,
          error: 'HTTP Error ' + responseCode,
          url: imageUrl,
          details: response.getContentText()
        };
      }

    } catch (e) {
      Logger.log('Excepción en generateImage: ' + e.toString());
      return {
        success: false,
        error: e.toString(),
        message: 'Error generando imagen con Pollinations.ai'
      };
    }
  }

  /**
   * Genera una imagen y la guarda en Google Drive
   * @param {string} prompt - Descripción de la imagen
   * @param {object} options - Opciones de generación
   * @param {string} folderId - ID de la carpeta de Drive (opcional)
   * @return {object} Resultado con información del archivo
   */
  function generateImageToDrive(prompt, options, folderId) {
    try {
      Logger.log('Generando imagen para guardar en Drive...');

      var result = generateImage(prompt, options);

      if (!result.success) {
        return result;
      }

      // Obtener carpeta
      var folder;
      if (folderId) {
        folder = DriveApp.getFolderById(folderId);
      } else {
        folder = DriveApp.getRootFolder();
      }

      // Crear nombre de archivo
      var timestamp = new Date().getTime();
      var filename = 'pollinations_' + timestamp + '.png';

      // Guardar blob en Drive
      var file = folder.createFile(result.blob);
      file.setName(filename);

      Logger.log('Imagen guardada en Drive: ' + file.getName());
      Logger.log('URL del archivo: ' + file.getUrl());

      return {
        success: true,
        fileId: file.getId(),
        fileName: filename,
        fileUrl: file.getUrl(),
        viewUrl: 'https://drive.google.com/file/d/' + file.getId() + '/view',
        downloadUrl: 'https://drive.google.com/uc?id=' + file.getId() + '&export=download',
        originalUrl: result.url,
        prompt: prompt,
        model: result.model,
        dimensions: result.dimensions
      };

    } catch (e) {
      Logger.log('Error guardando en Drive: ' + e.toString());
      return {
        success: false,
        error: e.toString(),
        message: 'Error guardando imagen en Drive'
      };
    }
  }

  /**
   * Prueba de conectividad con Pollinations.ai
   * @return {object} Estado de la conexión
   */
  function healthCheck() {
    try {
      Logger.log('Verificando conexión con Pollinations.ai...');

      var testPrompt = 'test';
      var testUrl = BASE_URL + '/' + encodeURIComponent(testPrompt) +
                    '?width=100&height=100&nologo=true';

      var response = UrlFetchApp.fetch(testUrl, {
        method: 'get',
        muteHttpExceptions: true
      });

      var code = response.getResponseCode();

      if (code === 200) {
        Logger.log('Conexión exitosa con Pollinations.ai');
        return {
          success: true,
          status: 'OK',
          message: 'Pollinations.ai está disponible',
          service: 'Pollinations.ai',
          requiresAuth: false,
          cost: 'Free'
        };
      } else {
        Logger.log('Error en health check: HTTP ' + code);
        return {
          success: false,
          status: 'ERROR',
          message: 'HTTP Error ' + code,
          service: 'Pollinations.ai'
        };
      }

    } catch (e) {
      Logger.log('Excepción en healthCheck: ' + e.toString());
      return {
        success: false,
        status: 'ERROR',
        message: e.toString(),
        service: 'Pollinations.ai'
      };
    }
  }

  /**
   * Función de prueba completa para Pollinations.ai
   */
  function testImageGeneration() {
    try {
      Logger.log('=== INICIANDO PRUEBA DE POLLINATIONS.AI ===');
      Logger.log('Este servicio es GRATUITO y no requiere API Key\n');

      // Test 1: Health Check
      Logger.log('1. Verificando conexión...');
      var health = healthCheck();
      Logger.log('   Estado: ' + (health.success ? 'OK ✓' : 'ERROR ✗'));
      Logger.log('   Mensaje: ' + health.message);

      if (!health.success) {
        Logger.log('\n⚠️ No se pudo conectar con Pollinations.ai');
        Logger.log('   Verifica tu conexión a internet');
        return { success: false, error: 'Health check failed' };
      }

      // Test 2: Generar imagen simple
      Logger.log('\n2. Generando imagen de prueba...');
      var testPrompt = 'A red sports car in a modern garage, photorealistic, high quality, 8k';
      Logger.log('   Prompt: ' + testPrompt);

      var result = generateImage(testPrompt, {
        width: 512,
        height: 512,
        nologo: true
      });

      if (result.success) {
        Logger.log('   ✓ Imagen generada exitosamente');
        Logger.log('   URL: ' + result.url);
        Logger.log('   Modelo: ' + result.model);
        Logger.log('   Dimensiones: ' + result.dimensions);
        Logger.log('   Tamaño: ' + Math.round(result.size / 1024) + ' KB');

        Logger.log('\n=== PRUEBA EXITOSA ===');
        Logger.log('Puedes usar generateImage() para generar más imágenes');
        Logger.log('O usa generateImageToDrive() para guardarlas en Google Drive');

        return {
          success: true,
          message: 'Prueba completada exitosamente',
          imageUrl: result.url,
          details: result
        };
      } else {
        Logger.log('   ✗ Error generando imagen: ' + result.error);
        if (result.details) {
          Logger.log('   Detalles: ' + result.details);
        }

        Logger.log('\n=== PRUEBA FALLIDA ===');
        return {
          success: false,
          error: result.error,
          details: result.details
        };
      }

    } catch (e) {
      Logger.log('Excepción en testImageGeneration: ' + e.toString());
      return {
        success: false,
        error: e.toString()
      };
    }
  }

  /**
   * Lista de modelos disponibles en Pollinations.ai
   * @return {array} Lista de modelos
   */
  function getAvailableModels() {
    return [
      { name: 'flux', description: 'Flux.1 - Alta calidad, rápido' },
      { name: 'flux-realism', description: 'Flux Realism - Fotorealista' },
      { name: 'any-dark', description: 'AnyDark - Estilo oscuro' },
      { name: 'stable-diffusion-3.5-large', description: 'SD 3.5 Large - Máxima calidad' },
      { name: 'stable-diffusion-3.5-medium', description: 'SD 3.5 Medium - Balanceado' },
      { name: 'stable-diffusion-xl', description: 'SDXL - Excelente calidad' },
      { name: 'turbo', description: 'Turbo - Más rápido' }
    ];
  }

  // Exportar funciones públicas
  return {
    generateImage: generateImage,
    generateImageToDrive: generateImageToDrive,
    healthCheck: healthCheck,
    testImageGeneration: testImageGeneration,
    getAvailableModels: getAvailableModels
  };

})();