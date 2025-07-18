import React, { useState } from 'react';
import axios from 'axios';

function EditModal({ slides, filename, onClose }) {
  const [editedSlides, setEditedSlides] = useState(
    [...slides]
      .sort((a, b) => a.id - b.id)
      .map(slide => ({
        ...slide,
        elements: slide.elements.map(el => ({ ...el }))
      }))
  );
  const [slideJsons, setSlideJsons] = useState(
    [...slides]
      .sort((a, b) => a.id - b.id)
      .map(slide => {
        const cleanSlide = {
          ...slide,
          elements: slide.elements.map(el => {
            if (el.type === 'image') {
              return { ...el, src: el.src.length > 50 ? '[IMAGE_URL]' : el.src };
            }
            return el;
          })
        };
        return JSON.stringify(cleanSlide, null, 2);
      })
  );
  const [imageFiles, setImageFiles] = useState({});
  const [downloading, setDownloading] = useState(false);
  const [jsonErrors, setJsonErrors] = useState({});
  const [selectedImageId, setSelectedImageId] = useState(null);
  const [uploadedImageName, setUploadedImageName] = useState('');

  const handleJsonChange = (slideIndex, value) => {
    const newJsons = [...slideJsons];
    newJsons[slideIndex] = value;
    setSlideJsons(newJsons);
    
    try {
      const parsedSlide = JSON.parse(value);
      const updated = [...editedSlides];
      updated[slideIndex] = parsedSlide;
      setEditedSlides(updated);
      
      const errors = { ...jsonErrors };
      delete errors[slideIndex];
      setJsonErrors(errors);
    } catch (error) {
      setJsonErrors({ ...jsonErrors, [slideIndex]: error.message });
    }
  };

  const handleImageUpload = async (file) => {
    if (!file) return;
    
    try {
      // Create form data for file upload
      const formData = new FormData();
      formData.append('image', file);
      
      // Upload the image to the server
      const response = await axios.post('/api/upload-image', formData, {
        headers: { 'Content-Type': 'multipart/form-data' }
      });
      
      // Get the image name from the response
      const imageName = response.data.imageName;
      setUploadedImageName(imageName);
      
      // Show confirmation
      console.log(`Image uploaded: ${imageName}\n\nYou can now use this name in your slide JSON.`);
    } catch (error) {
      alert('Error uploading image: ' + error.message);
    }
  };
  
  const handleClearUploads = async () => {
    try {
      const response = await axios.post('/api/clear-uploads');
      alert('Upload folder cleared successfully!');
      setUploadedImageName('');
      setImageFiles({});
    } catch (error) {
      alert('Error clearing uploads: ' + error.message);
    }
  };


  const handleEditData = async (slideIndex) => {
    try {
      const currentContent = slideJsons[slideIndex];
      const response = await axios.post('/api/edit-data', {
        content: JSON.parse(currentContent)
      });

      const updatedSlide = response.data.updatedData;
      const updated = [...editedSlides];
      updated[slideIndex] = updatedSlide;
      setEditedSlides(updated);
      
      const displaySlide = {
        ...updatedSlide,
        elements: updatedSlide.elements.map(el => 
          el.type === 'image' ? { ...el, src: el.src.length > 50 ? '[IMAGE_URL]' : el.src } : el
        )
      };
      
      const newJsons = [...slideJsons];
      newJsons[slideIndex] = JSON.stringify(displaySlide, null, 2);
      setSlideJsons(newJsons);
      
    } catch (error) {
      alert('Error editing data: ' + error.message);
    }
  };

  const handleDownload = async () => {
    // Check for JSON errors
    if (Object.keys(jsonErrors).length > 0) {
      alert('Please fix JSON errors before downloading');
      return;
    }
    
    setDownloading(true);
    
    try {
      const formData = new FormData();
      formData.append('slides', JSON.stringify(editedSlides));
      formData.append('filename', filename);
      
      Object.entries(imageFiles).forEach(([key, file]) => {
        formData.append(key, file);
      });

      const response = await axios.post('/api/generate', formData, {
        headers: { 'Content-Type': 'multipart/form-data' },
        responseType: 'blob'
      });

      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', 'updated-presentation.pptx');
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);
      
    } catch (error) {
      alert('Error generating file: ' + error.message);
    } finally {
      setDownloading(false);
    }
  };



  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center p-4 z-50">
      <div className="bg-white rounded-lg max-w-7xl w-full h-[95vh] flex flex-col">
        <div className="p-4 border-b sticky top-0 bg-white z-10">
          <div className="flex justify-between items-center">
            <h2 className="text-xl font-bold">Edit Presentation</h2>
            <div className="flex items-center space-x-2">
              <button
                onClick={handleClearUploads}
                className="bg-red-500 text-white px-3 py-1 rounded text-sm hover:bg-red-600"
              >
                Clear Uploads
              </button>
              <label className="bg-green-600 text-white px-3 py-1 rounded text-sm cursor-pointer hover:bg-green-700">
                Upload Image
                <input
                  type="file"
                  accept="image/*"
                  onChange={(e) => handleImageUpload(e.target.files[0])}
                  className="hidden"
                />
              </label>
              <button onClick={onClose} className="text-gray-500 hover:text-gray-700 text-2xl">×</button>
            </div>
          </div>
          {uploadedImageName && (
            <div className="mt-2 p-2 bg-green-50 border border-green-200 rounded text-sm">
              <span className="font-medium">Uploaded image:</span> {uploadedImageName}
            </div>
          )}
          <p className="text-sm text-gray-600 mt-2">
            <span className="font-medium">Instructions:</span> Upload images using the button above • Edit slide JSON directly in textareas below
          </p>
        </div>

        <div className="p-6 space-y-6 overflow-y-auto flex-grow">
          {/* Log slide IDs for debugging */}
          {console.log('Original slide IDs:', slides.map(s => s.id))}
          {console.log('Edited slide IDs:', editedSlides.map(s => s.id))}
          
          {/* Use numeric sort by ID to ensure correct sequence */}
          {[...slides].sort((a, b) => a.id - b.id).map((originalSlide) => {
            const slide = editedSlides.find(s => s.id === originalSlide.id);
            const slideIndex = editedSlides.findIndex(s => s.id === originalSlide.id);
            if (!slide) {
              console.log(`Missing slide with ID ${originalSlide.id}`);
              return null;
            }
            return (
            <div key={slide.id} className="border rounded-lg p-4">
              <div className="flex justify-between items-center mb-4">
                <h3 className="font-semibold">Slide {slide.id} JSON</h3>
                <span className="text-sm bg-blue-100 px-2 py-1 rounded-full">
                  Slide #{slide.id}
                </span>
              </div>
              
              <textarea
                value={slideJsons[slideIndex]}
                onChange={(e) => handleJsonChange(slideIndex, e.target.value)}
                className={`w-full h-96 p-3 border rounded-md font-mono text-sm ${
                  jsonErrors[slideIndex] ? 'border-red-500 bg-red-50' : 'border-gray-300'
                }`}
                placeholder="Edit slide JSON here..."
              />
              
              {jsonErrors[slideIndex] && (
                <div className="mt-2 p-2 bg-red-100 border border-red-300 rounded text-red-700 text-sm">
                  <strong>JSON Error:</strong> {jsonErrors[slideIndex]}
                </div>
              )}
              
              <button
                onClick={() => handleEditData(slideIndex)}
                className="mt-3 px-4 py-2 bg-blue-500 text-white rounded-md hover:bg-blue-600 text-sm"
              >
                Edit Data
              </button>
              
              <div className="mt-2 text-xs text-gray-600">
                <strong>Tip:</strong> Edit the JSON directly to modify slide elements. 
                Use the "Upload Image" button at the top to add images, then reference them by name in your JSON.
              </div>
            </div>
          );
          })}
        </div>

        <div className="p-6 border-t bg-gray-50 sticky bottom-0">
          <div className="flex justify-end space-x-4">
            <button
              onClick={onClose}
              className="px-4 py-2 text-gray-600 border border-gray-300 rounded-md hover:bg-gray-50"
            >
              Cancel
            </button>
            <button
              onClick={handleDownload}
              disabled={downloading}
              className="px-4 py-2 bg-green-600 text-white rounded-md hover:bg-green-700 disabled:bg-gray-400"
            >
              {downloading ? 'Generating...' : 'Download Updated PPTX'}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

export default EditModal;