import { Capacitor } from '@capacitor/core';

export type DocKind = 'pdf' | 'docx' | 'xlsx';

export interface OpenedFile {
  name: string;
  kind: DocKind;
  data: ArrayBuffer;
  /** Đường dẫn trên máy tính (chỉ có khi chạy Electron). */
  path?: string;
}

interface DesktopBridge {
  openFile(): Promise<{ name: string; path: string; data: ArrayBuffer } | null>;
  saveFile(opts: { name: string; path?: string; data: ArrayBuffer; saveAs?: boolean }): Promise<string | null>;
}

declare global {
  interface Window {
    dvdDesktop?: DesktopBridge;
  }
}

export const MIME: Record<DocKind, string> = {
  pdf: 'application/pdf',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

export const ACCEPT = '.pdf,.docx,.xlsx,' + Object.values(MIME).join(',');

export function kindFromName(name: string): DocKind | null {
  const ext = name.toLowerCase().split('.').pop();
  if (ext === 'pdf') return 'pdf';
  if (ext === 'docx') return 'docx';
  if (ext === 'xlsx') return 'xlsx';
  return null;
}

export function platform(): 'desktop' | 'ios' | 'android' | 'web' {
  if (window.dvdDesktop) return 'desktop';
  const p = Capacitor.getPlatform();
  if (p === 'ios' || p === 'android') return p;
  return 'web';
}

export async function fileToOpened(file: File): Promise<OpenedFile> {
  const kind = kindFromName(file.name);
  if (!kind) throw new Error(`Định dạng không hỗ trợ: ${file.name}`);
  return { name: file.name, kind, data: await file.arrayBuffer() };
}

/** Mở hộp thoại chọn tệp. Trên máy tính dùng hộp thoại hệ điều hành; nơi khác dùng input file. */
export async function pickFile(): Promise<OpenedFile | null> {
  if (window.dvdDesktop) {
    const res = await window.dvdDesktop.openFile();
    if (!res) return null;
    const kind = kindFromName(res.name);
    if (!kind) throw new Error(`Định dạng không hỗ trợ: ${res.name}`);
    return { ...res, kind };
  }
  return new Promise((resolve, reject) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = ACCEPT;
    input.onchange = () => {
      const f = input.files?.[0];
      if (!f) return resolve(null);
      fileToOpened(f).then(resolve, reject);
    };
    input.click();
  });
}

function toBase64(buf: ArrayBuffer): string {
  const bytes = new Uint8Array(buf);
  let bin = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    bin += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(bin);
}

/**
 * Lưu tệp theo nền tảng:
 * - Máy tính (Electron): ghi đè tệp gốc, hoặc hộp thoại "Lưu thành".
 * - iOS/Android: ghi vào thư mục Documents của ứng dụng rồi mở bảng Chia sẻ.
 * - Web: hộp thoại lưu của trình duyệt (nếu có) hoặc tải xuống.
 * Trả về đường dẫn mới (nếu có) để lần lưu sau ghi đè đúng chỗ.
 */
export async function saveFile(
  file: { name: string; kind: DocKind; path?: string },
  data: ArrayBuffer,
  saveAs = false,
): Promise<string | undefined> {
  if (window.dvdDesktop) {
    return (await window.dvdDesktop.saveFile({ name: file.name, path: file.path, data, saveAs })) ?? file.path;
  }
  if (Capacitor.isNativePlatform()) {
    const { Filesystem, Directory } = await import('@capacitor/filesystem');
    const { Share } = await import('@capacitor/share');
    const res = await Filesystem.writeFile({
      path: file.name,
      data: toBase64(data),
      directory: Directory.Documents,
    });
    try {
      await Share.share({ title: file.name, url: res.uri });
    } catch {
      /* người dùng đóng bảng chia sẻ: tệp vẫn đã được lưu */
    }
    return undefined;
  }
  const blob = new Blob([data], { type: MIME[file.kind] });
  const picker = (window as unknown as { showSaveFilePicker?: Function }).showSaveFilePicker;
  if (picker) {
    try {
      const handle = await picker({ suggestedName: file.name });
      const w = await handle.createWritable();
      await w.write(blob);
      await w.close();
      return undefined;
    } catch (e) {
      if ((e as DOMException).name === 'AbortError') return undefined;
    }
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = file.name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
  return undefined;
}
