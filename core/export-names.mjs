const invalidCharacters = /[<>:"/\\|?*\u0000-\u001f\u007f-\u009f]/;
const reservedName = /^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9]|CONIN\$|CONOUT\$)(?:\.|$)/i;
export const stripOutputExtension = (name) => name.replace(/\.(?:papertemplate|paperlink|flowtemplate|flowdesk|docx|pdf|png|jpe?g|html?|txt|xlsx|xlsm?|ods|csv|tsv|zip)$/i, '');
export function validateFileName(value) {
  const name = stripOutputExtension(String(value).trim()).normalize('NFC');
  if (!name || name === '.' || name === '..' || invalidCharacters.test(name) || /[. ]$/.test(name) || reservedName.test(name)) throw new Error('Enter a file name without a folder path or these characters: < > : " / \\ | ? *');
  return name;
}
export function suggestedFileName(value, fallback) {
  let name = stripOutputExtension(String(value).trim()).normalize('NFC').replace(new RegExp(invalidCharacters.source, 'g'), '_').replace(/[. ]+$/, '');
  if (!name) name = fallback;
  if (reservedName.test(name)) name = '_' + name;
  return [...name].slice(0, 160).join('').replace(/[. ]+$/, '') || fallback;
}
// ZIPs permit repeated or case-sensitive paths, but Windows extraction does not.
// Resolve collisions before writing any entry, including user-edited names.
export function uniqueFileNames(entries) {
  const used = new Set();
  return entries.map((entry) => {
    const stem = validateFileName(entry.name);
    let name = stem, count = 1;
    while (used.has(name.normalize('NFC').toLowerCase())) {
      const suffix = ' (' + (++count) + ')';
      name = [...stem].slice(0, 160 - suffix.length).join('').replace(/[. ]+$/, '') + suffix;
    }
    used.add(name.normalize('NFC').toLowerCase());
    return { ...entry, name };
  });
}
