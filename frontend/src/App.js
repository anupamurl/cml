import React, { useState } from 'react';
import axios from 'axios';
import EditModal from './EditModal';

function App() {
  const [file, setFile] = useState(null);
  const [slides, setSlides] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [loading, setLoading] = useState(false);
  const [filename, setFilename] = useState('');

  const handleFileChange = (e) => {
    setFile(e.target.files[0]);
  };

  const handleUpload = async () => {
    if (!file) return;

    setLoading(true);
    const formData = new FormData();
    formData.append('pptx', file);

    try {
      const response = await axios.post('/api/upload', formData, {
        headers: { 'Content-Type': 'multipart/form-data' }
      });
      
      setSlides(response.data.slides);
      setFilename(response.data.filename);
      setShowModal(true);
    } catch (error) {
      alert('Error uploading file: ' + error.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-100 py-12 px-4">
      <div className="max-w-md mx-auto bg-white rounded-lg shadow-md p-6">
        <h1 className="text-2xl font-bold text-center mb-6">PPTX Editor</h1>
        
        <div className="space-y-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Upload PowerPoint File
            </label>
            <input
              type="file"
              accept=".pptx"
              onChange={handleFileChange}
              className="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
            />
          </div>
          
          <button
            onClick={handleUpload}
            disabled={!file || loading}
            className="w-full bg-blue-600 text-white py-2 px-4 rounded-md hover:bg-blue-700 disabled:bg-gray-400 disabled:cursor-not-allowed"
          >
            {loading ? 'Processing...' : 'Upload & Edit'}
          </button>
        </div>
      </div>

      {showModal && (
        <EditModal
          slides={slides}
          filename={filename}
          onClose={() => setShowModal(false)}
        />
      )}
    </div>
  );
}

export default App;