// Bambu Studio runs this after every slice (Process > Others > Post-processing scripts).
// It notes the print time and filament use of the plate that was just sliced, so that
// "Send to Pricer" can open it later.
//
// It is run by wscript.exe, which shows no window, and is written in the old JScript that
// wscript understands. It never changes the G-code and always finishes without an error,
// because Bambu Studio fails the slice when a post-processing script reports one.

try {
  (function () {
    var args = WScript.Arguments;
    if (!args.length) return;
    var gcode = args(args.length - 1); // Bambu Studio adds the G-code path as the last argument.

    var fso = new ActiveXObject('Scripting.FileSystemObject');
    var shell = new ActiveXObject('WScript.Shell');
    if (!fso.FileExists(gcode)) return;

    // The numbers we need are comments in the header and settings blocks at the top of the file.
    var wanted = /^; (BambuStudio |model printing time|estimated printing time|total filament weight|filament: |filament_type =|filament_colour =)/;
    var kept = [];
    var input = fso.OpenTextFile(gcode, 1, false, 0);
    for (var n = 0; n < 5000 && !input.AtEndOfStream; n++) {
      var line = input.ReadLine();
      if (wanted.test(line)) kept.push(line);
      if (line.indexOf('; CONFIG_BLOCK_END') === 0 || line.indexOf('; EXECUTABLE_BLOCK_START') === 0) break;
    }
    input.Close();

    var folder = shell.ExpandEnvironmentStrings('%LOCALAPPDATA%') + '\\PrintPricer\\slices';
    makeFolder(fso, folder);

    // One note per plate: slicing the same plate again replaces its note. The end of the G-code path
    // is what tells plates apart, and a short name keeps the note's own path within Windows' limit.
    var name = gcode.replace(/[^A-Za-z0-9]+/g, '_').slice(-80) + '.txt';
    var out = fso.CreateTextFile(folder + '\\' + name, true, true);
    out.WriteLine('path=' + gcode);
    out.WriteLine('time=' + Math.floor(new Date().getTime() / 1000));
    for (var i = 0; i < kept.length; i++) out.WriteLine(kept[i]);
    out.Close();
  })();
} catch (e) {
  // Deliberately ignored: pricing is a convenience and must never get in the way of slicing.
}
WScript.Quit(0);

function makeFolder(fso, path) {
  if (fso.FolderExists(path)) return;
  makeFolder(fso, fso.GetParentFolderName(path));
  fso.CreateFolder(path);
}
