import { zipSync } from 'fflate';

type ZipFile = { path: string; data: Uint8Array };

export function buildZip(files: ZipFile[]) {
  const entries: Record<string, Uint8Array> = {};
  for (const file of files) {
    const path = file.path.replace(/\\/g, '/');
    entries[path] = file.data;
  }
  return zipSync(entries);
}
