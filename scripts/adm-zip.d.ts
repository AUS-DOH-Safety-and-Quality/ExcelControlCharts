declare module "adm-zip" {
  class ZipEntry {
    entryName: string;
    comment: string;
    attr: number;
    getData(): Buffer;
  }

  export default class AdmZip {
    constructor(input?: Buffer | string);
    getEntry(entryName: string): ZipEntry | null;
    getEntries(): ZipEntry[];
    readAsText(entry: ZipEntry): string;
    addFile(entryName: string, content: Buffer, comment?: string, attr?: number): void;
    writeZipPromise(targetFileName: string): Promise<void>;
  }

  export { ZipEntry as IZipEntry };
}