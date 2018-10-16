/**
 * Test script and library usage examples are provided here. 
 */
const { listAllDevices, WIADevice } = require('wia-scanner-js');

/**
 * listAllDevices - List all WIA devices available on the current computer.
 * returns = 
 * [
 *   {
 *     deviceUUID: L"{6BDD1FC6-810F-11D0-BEC7-08002BE2092F}\\0000",	
 *     manufacturer: L"Kodak",
 *     description: L"KODAK i2600 Scanner",
 * 	   deviceType: 65537,
 *     portName: "\\\\.\\Usbscan0",
 *     deviceName: "KODAK i2600 Scanner",
 *     serverName: "local",
 *     remoteDeviceId: "",
 *     uiCLSID: "{4DB1AD10-3391-11D2-9A33-00C04FA36145}",
 * 	   hardwareConfig: 0,
 *     baudrate: "",
 * 	   STIGenericCapabilities: 49,
 *     wiaVersion: "2.0",
 *     driverVersion: "2.2.0.0",
 *     PnPIDString: "\\\\?\\usb#vid_040a&pid_601d#0000000000000000#{6bdd1fc6-810f-11d0-bec7-08002be2092f}",
 * 	   STIDriverVersion: 3
 *   },
 *   ...
 * ]
 */
let devices = listAllDevices();

/**
 * WIADevice(deviceUUID) - Open a WIA device.
 *   deviceUUID: UUID of the device acquired from the method listAllDevices.
 * 
 * Throws error if the open action fails.
 * 
 * returns = WIADevice
 */
let wiaDevice = new WIADevice(devices[0].deviceUUID);

/**
 * wiaDevice.showDeviceDlg - Open the scan dialog provided by Microsoft Windows. Then enter the modal loop which quits as the dialog closes.
 * 
 * CAUTION - properties set by other methods of this WIADevice WILL NOT AFFECT settings showed in the system dialog. They are totally independent.
 * 
 * wiaDevice.showDeviceDlg(parentWindow, saveFolder, saveFilename)
 *   parentWindow: HWND handle of the parent window. 
 *                 For Example, the handle can be acquired by calling win.getNativeWindowHandle() when the current Node.js runtime is Electron.
 *                 In other circumstances, the parameter must be null.
 *   saveFolder:   Where to save images acquired from the scanner.
 *   saveFilename: Filename template of image files.
 * 
 * returns = 
 * {
 *   files: [
 *     "C:\\Users\\example\\Pictures\\scanner-test\\scan111.jpg",
 *     "C:\\Users\\example\\Pictures\\scanner-test\\scan111 (2).jpg",
 *     "C:\\Users\\example\\Pictures\\scanner-test\\scan111 (3).jpg",
 *     ...
 *   ]
 *   retCode: Number, // error code
 *   errorMsg: string // error msg
 * }
 */
let scanDialogRet = wiaDevice.showDeviceDlg(null, "C:\\Users\\example\\Pictures\\scanner-test", "scan111");

/**
 * wiaDevice.getSources() - list all available image sources from the WIA device.
 * 
 * returns = {
 *   sources: [
 *     "0000\Root\Feeder",
 *     "0000\Root\Feeder\Back",
 *     "0000\Root\Feeder\Front"
 *   ]
 * }
 */
let imageSourcesArray = wiaDevice.getSources();

/**
 * wiaDevice.getProperties() - get properties of the WIA device currently opened.
 * 
 * returns = {
 *   format: "tiff",                        // Output image format(tiff/bmp/jpeg/png)
 *   paper: "auto",                         // Paper type(auto/letter/businesscard/uslegal/usstatement/a0~a10/b0~b10)
 *   color: "blackwhite",                   // Color mode(blackwhite/greyscale/fullcolor)
 *   brightness_range: {min: -50, max: 50}, // Brightness value range available for this scanner
 *   brightness: 0,                         // Current brightness value
 *   contrast_range: {min: -50, max 50},    // Contrast value range available for this scanner
 *   contrast: 0,                           // Current contrast value
 *   dpi: 200,                              // Current DPI value
 *   // The following properties has significance only if the scanner is the feeder type
 *   pageCount: "all",                      // How many pages will be scanned("all"/ page number).
 *                                          // The scanner will keep running until no paper available if the value is set to "all"
 *   document_handling: "front"             // Which side of the paper will be scanned（front: front side only/duplex: both sides）
 * }
 */
let properties = wiaDevice.getProperties();

/**
 * wiaDevice.setProperties(params) - set properties of the WIA device currently opened.
 * 
 * // Same definition as getProperties
 * // partly modification is accepted
 * params = {
 *  format: 'jpeg',
 *  paper: 'letter',
 *  color: 'fullcolor',
 *  brightness: 20,
 *  contrast: 50,
 *  dpi: 300,
 *  pageCount: 'all',
 *  document_handling: 'front',
 * }
 * 
 * returns = {
 *   // returns whether the modification is successful or not
 *   format: true,
 *   paper: true,
 *   ...
 * }
 */
let setPropsRet = wiaDevice.setProperties({
    format: 'jpeg',
    paper: 'letter',
    color: 'fullcolor',
    brightness: 20,
    contrast: 50,
    dpi: 300,
    // The following properties has significance only if the scanner is the feeder type
    pageCount: 'all',
    document_handling: 'front',
});

/**
 * event 'progress' - Indicates the progress of the current scan operation.
 * 
 * progressInfo = {
 *   page: 0,             // Current page index
 *   percent: 100,        // Scan progress of current page
 *   bytesTrnasferred: 0, // how many bytes has been transferred
 *   error: 0,            // Error code
 *   errMsg: ""           // Error msg
 * }
 * 
 */
wiaDevice.on('progress', (progressInfo) => {
    console.log(`page=${progressInfo.page}  percent=${progressInfo.percent}`);
});

/**
 * event 'complete' - Triggered after the scan operation has completed
 * 
 * imageData = {
 *   retCode: 0,          // Return value of the scan operation
 *   errMsg: "",          // Error msg
 *   files: [             // An array of the path of the acquired images
 *     "C:\\Users\\example\\Pictures\\scanner-test\\scan111_1.jpeg",
 *     "C:\\Users\\example\\Pictures\\scanner-test\\scan111_2.jpeg",
 *     "C:\\Users\\example\\Pictures\\scanner-test\\scan111_3.jpeg",
 *     ...
 *   ]
 * }
 * 
 */
wiaDevice.on('complete', (imageData) => {
    console.log(imageData);
});

/**
 * wiaDevice.getPreview(callback) - Get a preview image from the scanner
 * 
 * Currently not implemented yet.
 * 
 * imageData = {
 *   data: <dataURL>     // dataURL of the picture
 * }
 */
wiaDevice.getPreview(function (imageData) {

});

let isFeeder = wiaDevice.isFeeder();

/**
 * wiaDevice.doScan(params[, callback]) - Run the scan operation.
 * 
 * params = {
 *   saveDir: "C:\\Users\\example\\Pictures\\scanner-test", // Where to save images acquired from the scanner
 *   saveFilename: "test111"                                // Filename template of image files.
 * }
 * 
 * callback = function(imageData) {  // callback here will override the callback handling the event 'complete'!
 * 
 * }
 * 
 */
wiaDevice.doScan({
    saveDir: "C:\\Users\\example\\Pictures\\scanner-test",
    saveFilename: "test111"
});

/**
 * wiaDevice.cancel() - Abort the scan operation currently running.
 * 
 */
//setTimeout(() => {
//    wiaDevice.cancel()
//}, 2000);