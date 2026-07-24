const VehicleImageService = (() => {

  function normalize(text) {
    if (!text) return "";

    return String(text)
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/-+/g, "-")
      .replace(/^-|-$/g, "");
  }

  function buildImageName(vehicle) {
    return [
      normalize(vehicle.make),
      normalize(vehicle.model),
      normalize(vehicle.year),
      normalize(vehicle.color)
    ].filter(Boolean).join("-");
  }

  function buildImageFilename(vehicle) {
    return DriveService.buildImageFilename(buildImageName(vehicle));
  }

  function isImageFileCurrent(vehicle, fileId) {
    const file = DriveService.getFileById(fileId);

    if (!file) return false;

    return file.getName() === buildImageFilename(vehicle);
  }

  function ensureVehicleImage(vehicle) {
    const imageName = buildImageName(vehicle);

    if (!imageName) {
      return {
        success: false,
        error: "Vehicle image name could not be built."
      };
    }

    const existingFile = DriveService.findImage(imageName);

    if (existingFile) {
      return {
        success: true,
        fileId: existingFile.getId(),
        fileName: existingFile.getName(),
        source: "drive-cache"
      };
    }

    const prompt = PromptService.buildVehiclePrompt(vehicle);

    return PollinationsService.generateImageToDrive(
      prompt,
      {
        width: 1024,
        height: 1024,
        model: "flux",
        nologo: true,
        imageName: imageName
      }
    );
  }

  return Object.freeze({
    normalize,
    buildImageName,
    buildImageFilename,
    isImageFileCurrent,
    ensureVehicleImage
  });

})();