import { useState, useCallback } from 'react';

export function useFileSystem() {
  const [directoryHandle, setDirectoryHandle] = useState<FileSystemDirectoryHandle | null>(null);

  const requestPermission = useCallback(async () => {
    try {
      const handle = await (window as any).showDirectoryPicker({
        mode: 'readwrite',
      });
      setDirectoryHandle(handle);
      return handle;
    } catch (err) {
      console.error('Directory picker failed', err);
      return null;
    }
  }, []);

  const saveFile = useCallback(async (
    handle: FileSystemDirectoryHandle,
    blob: Blob,
    vendor: string,
    date: string,
    type: 'Statement' | 'Invoice'
  ) => {
    try {
      // Create subfolders: /YYYY/MM/
      const [year, month] = date.split('-');
      const yearFolder = await handle.getDirectoryHandle(year, { create: true });
      const monthFolder = await yearFolder.getDirectoryHandle(month, { create: true });

      const fileName = `${vendor}_${type}_${date}.pdf`.replace(/[/\\?%*:|"<>]/g, '-');
      const fileHandle = await monthFolder.getFileHandle(fileName, { create: true });
      
      const writable = await fileHandle.createWritable();
      await writable.write(blob);
      await writable.close();
      
      return `${year}/${month}/${fileName}`;
    } catch (err) {
      console.error('Failed to save file', err);
      throw err;
    }
  }, []);

  return {
    directoryHandle,
    requestPermission,
    saveFile
  };
}
