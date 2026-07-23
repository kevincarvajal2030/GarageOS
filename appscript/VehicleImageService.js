/**
 * ============================================================
 * GARAGE OS
 * VehicleImageService.gs
 *
 * Handles vehicle image naming and lookup logic.
 * ============================================================
 */

const VehicleImageService = (() => {

  /**
   * Normalizes text for filenames.
   *
   * Example:
   * "Land Rover"
   * ↓
   * "land-rover"
   */
  function normalize(text) {

    if (!text) return "";

    return String(text)
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ")
      .replace(/ /g, "-");

  }

  /**
   * Builds the image filename (without extension).
   *
   * Example:
   * tesla-model-y-2022-red
   */
  function buildImageName(vehicle) {

    return [

      normalize(vehicle.make),
      normalize(vehicle.model),
      normalize(vehicle.year),
      normalize(vehicle.color)

    ].join("-");

  }


  /**
 * Returns the existing image file for a vehicle.
 * Returns null if no image exists.
 */
  function findExistingImage(vehicle) {

    const imageName = buildImageName(vehicle);

    return DriveService.findImage(imageName);

  }

  return Object.freeze({

    buildImageName,
    findExistingImage

  });

})();


function testFindExistingVehicleImage() {

  const vehicle = {

    make: "GMC",
    model: "Sierra",
    year: 2020,
    color: "Black"

  };

  const file = VehicleImageService.findExistingImage(vehicle);

  if (!file) {

    Logger.log("Image not found.");
    return;

  }

  Logger.log(file.getName());
  Logger.log(file.getId());

}

