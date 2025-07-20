import React, { useState, useEffect } from 'react';
import axios from 'axios';
import EditModal from './EditModal';

function App() {
  const [file, setFile] = useState(null);
  const [slides, setSlides] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [loading, setLoading] = useState(false);
  const [filename, setFilename] = useState('');
  const [templateName, setTemplateName] = useState('');
  const [templates, setTemplates] = useState([]);
  const [clearingUploads, setClearingUploads] = useState(false);
  const [downloading, setDownloading] = useState(false);
  const [deletingTemplate, setDeletingTemplate] = useState(null);
  const [editingTemplate, setEditingTemplate] = useState(null);

  const handleFileChange = (e) => {
    setFile(e.target.files[0]);
  };
  
  // Handle template editing
  const handleEditTemplate = async (templateId) => {
    setEditingTemplate(templateId);
    try {
      const response = await axios.get(`/api/templates/${templateId}`);
      setSlides(response.data.slides);
      setFilename(''); // No filename for template editing
      setTemplateName(response.data.templateName);
      setShowModal(true);
    } catch (error) {
      alert('Error loading template: ' + error.message);
    } finally {
      setEditingTemplate(null);
    }
  };

  // Presentations fetching removed as we no longer display them
  
  // Fetch saved templates from MongoDB
  const fetchTemplates = async () => {
    try {
      const response = await axios.get('/api/templates');
      setTemplates(response.data.templates || []);
    } catch (error) {
      console.error('Error fetching templates:', error);
    }
  };
  
  // Handle template download
  const handleDownloadTemplate = async (templateId, templateName) => {
    setDownloading(templateId);
    try {
      const response = await axios.get(`/api/generate-template/${templateId}`, {
        responseType: 'blob'
      });
      
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', `${templateName}.pptx`);
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);
    } catch (error) {
      alert('Error downloading template: ' + error.message);
    } finally {
      setDownloading(null);
    }
  };
  
  // Handle template deletion
  const handleDeleteTemplate = async (templateId) => {
    if (!window.confirm('Are you sure you want to delete this template?')) return;
    
    setDeletingTemplate(templateId);
    try {
      await axios.delete(`/api/templates/${templateId}`);
      fetchTemplates(); // Refresh the list
    } catch (error) {
      alert('Error deleting template: ' + error.message);
    } finally {
      setDeletingTemplate(null);
    }
  };

  // Clear uploads folder
  const handleClearUploads = async () => {
    setClearingUploads(true);
    try {
      await axios.post('/api/clear-uploads');
      alert('Upload folder cleared successfully!');
    } catch (error) {
      alert('Error clearing uploads: ' + error.message);
    } finally {
      setClearingUploads(false);
    }
  };

  // Load templates on component mount
  useEffect(() => {
    fetchTemplates();
  }, []);

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
      setTemplateName(''); // Reset template name for new uploads
      setShowModal(true);
    } catch (error) {
      alert('Error uploading file: ' + error.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-100 py-6 px-4">
      {/* Header with Logo and Clear Uploads button */}
      <div className="max-w-6xl mx-auto mb-6 flex justify-between items-center">
        <div className="flex items-center">
          <div className="bg-blue-600 text-white p-2 rounded-md mr-3">
            <svg xmlns="http://www.w3.org/2000/svg" className="h-8 w-8" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
          </div>
          <h1 className="text-2xl font-bold">PPTX Editor</h1>
        </div>
        <button
          onClick={handleClearUploads}
          disabled={clearingUploads}
          className="bg-red-500 text-white px-4 py-2 rounded text-sm hover:bg-red-600 disabled:bg-red-300"
        >
          {clearingUploads ? 'Clearing...' : 'Clear Uploads'}
        </button>
      </div>
      
      <div className="max-w-6xl mx-auto grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Upload Form */}
        <div className="bg-white rounded-lg shadow-md p-6">
          <h2 className="text-xl font-semibold mb-4">Upload Presentation</h2>
          <div className="space-y-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Select PowerPoint File
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
        
        {/* Tables Container */}
        <div className="lg:col-span-2 space-y-6">
          {/* Saved Templates Table */}
          <div className="bg-white rounded-lg shadow-md p-6">
            <h2 className="text-xl font-semibold mb-4">Saved Templates</h2>
            
            {templates.length > 0 ? (
              <div className="overflow-x-auto">
                <table className="min-w-full divide-y divide-gray-200">
                  <thead className="bg-gray-50">
                    <tr>
                      <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Template Name</th>
                      <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Created Date</th>
                      <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Actions</th>
                    </tr>
                  </thead>
                  <tbody className="bg-white divide-y divide-gray-200">
                    {templates.map((template, index) => (
                      <tr key={template._id} className={index % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                        <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">{template.templateName}</td>
                        <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">{new Date(template.createdAt).toLocaleString()}</td>
                        <td className="px-6 py-4 whitespace-nowrap text-sm font-medium">
                          <button 
                            onClick={() => handleEditTemplate(template._id)}
                            disabled={editingTemplate === template._id}
                            className="text-blue-600 hover:text-blue-900 mr-3"
                          >
                            {editingTemplate === template._id ? 'Loading...' : 'Edit'}
                          </button>
                          <button 
                            onClick={() => handleDownloadTemplate(template._id, template.templateName)}
                            disabled={downloading === template._id}
                            className="text-green-600 hover:text-green-900 mr-3"
                          >
                            {downloading === template._id ? 'Downloading...' : 'Download'}
                          </button>
                          <button 
                            onClick={() => handleDeleteTemplate(template._id)}
                            disabled={deletingTemplate === template._id}
                            className="text-red-600 hover:text-red-900"
                          >
                            {deletingTemplate === template._id ? 'Deleting...' : 'Delete'}
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="text-center py-8 text-gray-500">No templates saved yet</div>
            )}
          </div>
        </div>
      </div>

      {showModal && (
        <EditModal
          slides={slides}
          filename={filename}
          initialTemplateName={templateName}
          onClose={() => {
            setShowModal(false);
            setTemplateName(''); // Reset template name
            fetchTemplates(); // Refresh templates after editing
          }}
        />
      )}
    </div>
  );
}

export default App;