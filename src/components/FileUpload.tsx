'use client';

import { useState } from 'react';

interface FileUploadProps {
  endpoint: string;
  label: string;
  accept?: string;
  extraFields?: Record<string, string>;
  onSuccess?: (data: unknown) => void;
}

export default function FileUpload({
  endpoint,
  label,
  accept = '.xlsx,.csv',
  extraFields,
  onSuccess,
}: FileUploadProps) {
  const [status, setStatus] = useState<'idle' | 'uploading' | 'success' | 'error'>('idle');
  const [message, setMessage] = useState('');

  const handleSubmit = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const form = e.currentTarget;
    const formData = new FormData(form);

    // Add extra fields
    if (extraFields) {
      for (const [key, value] of Object.entries(extraFields)) {
        formData.set(key, value);
      }
    }

    setStatus('uploading');
    setMessage('');

    try {
      const res = await fetch(endpoint, {
        method: 'POST',
        body: formData,
      });

      const data = await res.json();

      if (res.ok) {
        setStatus('success');
        setMessage(JSON.stringify(data, null, 2));
        onSuccess?.(data);
      } else {
        setStatus('error');
        setMessage(data.error || 'Import failed');
      }
    } catch (err) {
      setStatus('error');
      setMessage(err instanceof Error ? err.message : 'Network error');
    }
  };

  return (
    <form onSubmit={handleSubmit} className="space-y-3">
      <label className="block text-sm font-medium text-gray-700">{label}</label>
      <div className="flex gap-3 items-center">
        <input
          type="file"
          name="file"
          accept={accept}
          required
          className="text-sm text-gray-600 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-sm file:font-medium file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
        />
        <button
          type="submit"
          disabled={status === 'uploading'}
          className="px-4 py-2 bg-blue-600 text-white text-sm rounded-lg hover:bg-blue-700 disabled:bg-gray-400 disabled:cursor-not-allowed"
        >
          {status === 'uploading' ? 'Importando...' : 'Importa'}
        </button>
      </div>

      {status === 'success' && (
        <div className="bg-green-50 border border-green-200 rounded-lg p-3 text-sm text-green-800">
          <p className="font-medium">Import completato!</p>
          <pre className="mt-1 text-xs whitespace-pre-wrap">{message}</pre>
        </div>
      )}

      {status === 'error' && (
        <div className="bg-red-50 border border-red-200 rounded-lg p-3 text-sm text-red-800">
          <p className="font-medium">Errore</p>
          <p className="mt-1">{message}</p>
        </div>
      )}
    </form>
  );
}
