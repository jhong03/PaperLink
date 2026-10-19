// electron-builder 26.15.3 starts the NSIS helper with only __COMPAT_LAYER.
// Preserve Windows loader/temp settings when it supplies a partial environment.
// This preload affects packaging only and is not included in the distributed app.
if (process.platform === 'win32') {
  const { WineVmManager } = require('app-builder-lib/out/vm/WineVm.js');
  const original = WineVmManager.prototype.execWine;
  const windowsKeys = new Set(['SYSTEMROOT', 'WINDIR', 'PATH', 'TEMP', 'TMP', 'COMSPEC', 'PATHEXT', 'USERPROFILE', 'APPDATA', 'LOCALAPPDATA', 'PROGRAMDATA', 'PROGRAMFILES', 'PROGRAMFILES(X86)', 'COMMONPROGRAMFILES', 'COMMONPROGRAMFILES(X86)', 'SYSTEMDRIVE', 'HOMEDRIVE', 'HOMEPATH']);
  WineVmManager.prototype.execWine = async function (payload) {
    if (!payload.options?.env) return original.call(this, payload);
    const inherited = Object.fromEntries(Object.entries(process.env).filter(([key]) => windowsKeys.has(key.toUpperCase())));
    const options = { ...payload, options: { ...payload.options, env: { ...inherited, ...payload.options.env } } };
    for (let attempt = 0; ; attempt++) {
      try { return await original.call(this, options); }
      catch (failure) {
        // Windows can briefly refuse to start the freshly written NSIS helper.
        // Retry only a failed spawn, never a helper that started and then failed.
        if (payload.options.env.__COMPAT_LAYER !== 'RunAsInvoker' || failure.code !== 'UNKNOWN' || failure.syscall !== 'spawn' || attempt >= 4) throw failure;
        process.stderr.write('Waiting to retry the Windows packaging helper…\n');
        await new Promise((resolve) => setTimeout(resolve, 2000 * (attempt + 1)));
      }
    }
  };
}
