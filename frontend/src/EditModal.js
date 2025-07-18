import React, { useState } from 'react';
import axios from 'axios';

function EditModal({ slides, filename, onClose }) {
  const [editedSlides, setEditedSlides] = useState(
    slides.map(slide => ({
      ...slide,
      elements: slide.elements.map(el => ({ ...el }))
    }))
  );
  const [slideJsons, setSlideJsons] = useState(
    slides.map(slide => {
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

  const handleImageUpload = (slideIndex, file) => {
    if (!file) return;
    
    const key = `slide-${slideIndex}-new-image`;
    setImageFiles({ ...imageFiles, [key]: file });
    
    const reader = new FileReader();
    reader.onload = (e) => {
      const imageElement = {
        type: 'image',
        id: `image-${Date.now()}`,
        src: '[NEW_IMAGE]',
        x: 1,
        y: 1,
        width: 3,
        height: 2
      };
      
      const updated = [...editedSlides];
      updated[slideIndex].elements.push(imageElement);
      setEditedSlides(updated);
      
      const newJsons = [...slideJsons];
      const displaySlide = {
        ...updated[slideIndex],
        elements: updated[slideIndex].elements.map(el => 
          el.type === 'image' ? { ...el, src: el.src.length > 50 ? '[IMAGE_URL]' : el.src } : el
        )
      };
      newJsons[slideIndex] = JSON.stringify(displaySlide, null, 2);
      setSlideJsons(newJsons);
    };
    reader.readAsDataURL(file);
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
      <div className="bg-white rounded-lg max-w-7xl w-full max-h-[95vh] overflow-y-auto">
        <div className="p-4 border-b">
          <div className="flex justify-between items-center">
            <h2 className="text-xl font-bold">Edit Presentation</h2>
            <button onClick={onClose} className="text-gray-500 hover:text-gray-700 text-2xl">×</button>
          </div>
          <p className="text-sm text-gray-600 mt-1">
            <span className="font-medium">Instructions:</span> Edit slide JSON directly in textareas below • Use "Add Image" to upload new images
          </p>
        </div>

        <div className="p-6 space-y-6">
          {editedSlides.map((slide, slideIndex) => (
            <div key={slide.id} className="border rounded-lg p-4">
              <div className="flex justify-between items-center mb-4">
                <h3 className="font-semibold">Slide {slide.id} JSON</h3>
                <label className="bg-blue-500 text-white px-3 py-1 rounded text-sm cursor-pointer hover:bg-blue-600">
                  Add Image
                  <input
                    type="file"
                    accept="image/*"
                    onChange={(e) => handleImageUpload(slideIndex, e.target.files[0])}
                    className="hidden"
                  />
                </label>
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
                Use "Add Image" button to upload new images.
              </div>
            </div>
          ))}
        </div>

        <div className="p-6 border-t bg-gray-50">
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