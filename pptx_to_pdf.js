// PowerPoint PDF exporter for Windows Script Host.
//
// Usage:
//   cscript.exe //nologo pptx_to_pdf.js
//   cscript.exe //nologo pptx_to_pdf.js file1.pptx file2.pptx
//
// Each PDF is written next to its source presentation and is left in place.

(function () {
  "use strict";

  var PP_SAVE_AS_PDF = 32;
  var PP_ALERTS_NONE = 1;
  var MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3;
  var FSO = new ActiveXObject("Scripting.FileSystemObject");
  var SHELL = new ActiveXObject("WScript.Shell");
  var IS_CONSOLE = /cscript\.exe$/i.test(WScript.FullName);
  var files = collectInputFiles();
  var powerpoint = null;
  var failures = [];

  if (files.length === 0) {
    finish("PowerPoint file was not found in the current directory.", 2, true);
  }

  try {
    powerpoint = WScript.CreateObject("PowerPoint.Application");
    try {
      powerpoint.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE;
    } catch (securityError) {
      // Continue on older PowerPoint versions that do not expose this property.
    }
    try {
      powerpoint.DisplayAlerts = PP_ALERTS_NONE;
    } catch (alertError) {
      // Continue when this property cannot be changed.
    }

    for (var i = 0; i < files.length; i++) {
      try {
        exportPdf(powerpoint, files[i]);
      } catch (error) {
        failures.push(files[i] + "\n  " + formatError(error));
        log("ERROR: " + failures[failures.length - 1]);
      }
    }
  } catch (error) {
    finish("PowerPoint could not be started.\n" + formatError(error), 1, true);
  } finally {
    if (powerpoint !== null) {
      try {
        powerpoint.Quit();
      } catch (quitError) {
        // Ignore shutdown errors.
      }
    }
  }

  if (failures.length > 0) {
    finish(
      "Finished with errors: " + failures.length + " file(s).\n\n" + failures.join("\n\n"),
      1,
      true
    );
  }

  finish("PDF conversion completed: " + files.length + " file(s).", 0, false);

  function log(message) {
    if (IS_CONSOLE) {
      WScript.Echo(message);
    }
  }

  function finish(message, exitCode, isError) {
    if (IS_CONSOLE) {
      WScript.Echo(message);
    } else {
      // One dialog only after the whole batch. 16 is error icon, 64 information.
      SHELL.Popup(message, 0, "PowerPoint PDF conversion", isError ? 16 : 64);
    }
    WScript.Quit(exitCode);
  }

  function collectInputFiles() {
    var result = [];
    var seen = {};

    if (WScript.Arguments.Count() > 0) {
      for (var i = 0; i < WScript.Arguments.Count(); i++) {
        addPath(WScript.Arguments.Item(i), result, seen);
      }
    } else {
      addDirectory(SHELL.CurrentDirectory, result, seen);
    }

    result.sort(caseInsensitiveCompare);
    return result;
  }

  function addPath(path, result, seen) {
    var absolutePath = FSO.GetAbsolutePathName(path);

    if (FSO.FolderExists(absolutePath)) {
      addDirectory(absolutePath, result, seen);
      return;
    }

    if (!FSO.FileExists(absolutePath)) {
      log("SKIP: File not found: " + absolutePath);
      return;
    }

    if (!isPowerPointFile(absolutePath)) {
      log("SKIP: Unsupported file: " + absolutePath);
      return;
    }

    var key = absolutePath.toLowerCase();
    if (!seen[key]) {
      seen[key] = true;
      result.push(absolutePath);
    }
  }

  function addDirectory(directoryPath, result, seen) {
    var folder = FSO.GetFolder(directoryPath);
    var enumerator = new Enumerator(folder.Files);

    for (; !enumerator.atEnd(); enumerator.moveNext()) {
      var file = enumerator.item();
      if (isPowerPointFile(file.Path)) {
        addPath(file.Path, result, seen);
      }
    }
  }

  function isPowerPointFile(path) {
    var extension = FSO.GetExtensionName(path).toLowerCase();
    return extension === "pptx" || extension === "pptm";
  }

  function exportPdf(powerpoint, sourcePath) {
    var sourceFile = FSO.GetFile(sourcePath);
    var outputPath = FSO.BuildPath(
      sourceFile.ParentFolder.Path,
      FSO.GetBaseName(sourceFile.Name) + ".pdf"
    );
    var tempOutputPath = FSO.BuildPath(
      sourceFile.ParentFolder.Path,
      FSO.GetBaseName(sourceFile.Name) + ".__pptx_to_pdf_tmp__.pdf"
    );
    var presentation = null;

    log("Converting: " + sourceFile.Name);

    try {
      // ReadOnly=true, Untitled=false, WithWindow=false
      presentation = powerpoint.Presentations.Open(sourceFile.Path, true, false, false);

      if (FSO.FileExists(tempOutputPath)) {
        FSO.DeleteFile(tempOutputPath, true);
      }

      presentation.SaveAs(tempOutputPath, PP_SAVE_AS_PDF);

      if (!FSO.FileExists(tempOutputPath)) {
        throw new Error(
          "PowerPoint returned without creating the PDF: " + tempOutputPath
        );
      }

      if (FSO.FileExists(outputPath)) {
        FSO.DeleteFile(outputPath, true);
      }
      FSO.MoveFile(tempOutputPath, outputPath);

      log("  -> " + outputPath);
    } finally {
      if (presentation !== null) {
        try {
          presentation.Close();
        } catch (closeError) {
          // Ignore close errors and continue with the remaining files.
        }
      }
      if (FSO.FileExists(tempOutputPath)) {
        try {
          FSO.DeleteFile(tempOutputPath, true);
        } catch (cleanupError) {
          // Leave a diagnostic temp file only if Windows refuses deletion.
        }
      }
    }
  }

  function caseInsensitiveCompare(left, right) {
    left = left.toLowerCase();
    right = right.toLowerCase();
    if (left < right) {
      return -1;
    }
    if (left > right) {
      return 1;
    }
    return 0;
  }

  function formatError(error) {
    var number = "";
    var description = "";

    try {
      number = "0x" + (error.number >>> 0).toString(16);
    } catch (numberError) {
      number = "unknown";
    }

    try {
      description = error.description || error.message || String(error);
    } catch (descriptionError) {
      description = "Unknown error";
    }

    return number + ": " + description;
  }
}());
