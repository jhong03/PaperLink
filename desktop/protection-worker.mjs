import { parentPort } from 'node:worker_threads';
import { randomBytes } from 'node:crypto';
import { PDF } from '@libpdf/core';
import officeCrypto from 'officecrypto-tool';

parentPort.on('message', async (job) => {
  let input;
  try {
    input = Buffer.from(job.bytes);
    let output;
    if (job.format === 'pdf') {
      const pdf = await PDF.load(input);
      pdf.setProtection({ userPassword: job.password, ownerPassword: randomBytes(32).toString('hex'), algorithm: 'AES-256' });
      output = new Uint8Array(await pdf.save());
      const protectedDocument = await PDF.load(output);
      if (!protectedDocument.isEncrypted || protectedDocument.isAuthenticated) throw new Error('Protection was not applied.');
    } else {
      // The default is Office Agile encryption (AES-256 / SHA-512).
      // Passing type: 'standard' would select the older Office encryption.
      const encrypted = officeCrypto.encrypt(input, { password: job.password });
      if (!officeCrypto.isEncrypted(encrypted)) throw new Error('Protection was not applied.');
      output = new Uint8Array(encrypted);
    }
    parentPort.postMessage({ id: job.id, bytes: output }, [output.buffer]);
  } catch {
    parentPort.postMessage({ id: job.id, error: 'Could not password-protect this document. No unprotected attachment was prepared.' });
  } finally {
    input?.fill(0); job.bytes?.fill(0); job.password = '';
  }
});
