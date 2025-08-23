'use client';

import { useState, useEffect } from 'react';
import { Note, CreateNoteRequest, UpdateNoteRequest } from '@/types/note';
import TagInput from './TagInput';

interface NoteFormProps {
  note?: Note;
  onSubmit: (data: CreateNoteRequest | UpdateNoteRequest) => Promise<void> | void;
  onCancel: () => void;
  loading?: boolean;
}

export default function NoteForm({ note, onSubmit, onCancel, loading = false }: NoteFormProps) {
  const [title, setTitle] = useState('');
  const [content, setContent] = useState('');
  const [tags, setTags] = useState<string[]>([]);
  const [errors, setErrors] = useState<{ title?: string; content?: string }>({});

  useEffect(() => {
    if (note) {
      setTitle(note.title);
      setContent(note.content);
      setTags(note.tags || []);
    }
  }, [note]);

  const validateForm = () => {
    const newErrors: { title?: string; content?: string } = {};

    if (!title.trim()) {
      newErrors.title = 'Title is required';
    }

    if (!content.trim()) {
      newErrors.content = 'Content is required';
    }

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    
    if (!validateForm()) {
      return;
    }

    const formData = {
      title: title.trim(),
      content: content.trim(),
      tags: tags.join(';'), // Convert array to semicolon-separated string
    };

    onSubmit(formData);
  };

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center p-4 z-50">
      <div className="bg-white rounded-lg shadow-xl max-w-2xl w-full max-h-[90vh] overflow-y-auto">
        <div className="p-6">
          <h2 className="text-2xl font-bold text-gray-800 mb-6">
            {note ? 'Edit Note' : 'Create New Note'}
          </h2>

          <form onSubmit={handleSubmit} className="space-y-6">
            <div>
              <label htmlFor="title" className="block text-sm font-medium text-gray-700 mb-2">
                Title *
              </label>
              <input
                type="text"
                id="title"
                value={title}
                onChange={(e) => setTitle(e.target.value)}
                className={`w-full px-3 py-2 border rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 text-gray-900 bg-white ${
                  errors.title ? 'border-red-500' : 'border-gray-300'
                }`}
                placeholder="Enter note title"
                disabled={loading}
              />
              {errors.title && (
                <p className="mt-1 text-sm text-red-600">{errors.title}</p>
              )}
            </div>

            <div>
              <label htmlFor="content" className="block text-sm font-medium text-gray-700 mb-2">
                Content *
              </label>
              <textarea
                id="content"
                value={content}
                onChange={(e) => setContent(e.target.value)}
                rows={8}
                className={`w-full px-3 py-2 border rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 resize-vertical text-gray-900 bg-white ${
                  errors.content ? 'border-red-500' : 'border-gray-300'
                }`}
                placeholder="Enter note content"
                disabled={loading}
              />
              {errors.content && (
                <p className="mt-1 text-sm text-red-600">{errors.content}</p>
              )}
            </div>

            <div>
              <label htmlFor="tags" className="block text-sm font-medium text-gray-700 mb-2">
                Tags
              </label>
              <TagInput
                tags={tags}
                onChange={setTags}
                placeholder="Add tags..."
                disabled={loading}
              />
            </div>

            <div className="flex justify-end space-x-3 pt-4">
              <button
                type="button"
                onClick={onCancel}
                className="px-4 py-2 text-gray-700 bg-gray-100 rounded-md hover:bg-gray-200 transition-colors duration-200 disabled:opacity-50"
                disabled={loading}
              >
                Cancel
              </button>
              <button
                type="submit"
                className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors duration-200 disabled:opacity-50 disabled:cursor-not-allowed"
                disabled={loading}
              >
                {loading ? (
                  <div className="flex items-center space-x-2">
                    <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></div>
                    <span>{note ? 'Updating...' : 'Creating...'}</span>
                  </div>
                ) : (
                  note ? 'Update Note' : 'Create Note'
                )}
              </button>
            </div>
          </form>
        </div>
      </div>
    </div>
  );
}
