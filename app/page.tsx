'use client'

import { useState } from 'react'

export default function Home() {
  const [file, setFile] = useState<File | null>(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null)

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0]
    if (selectedFile) {
      // Validate file type
      if (!selectedFile.name.endsWith('.docx')) {
        setError('Please upload a .docx file only')
        setFile(null)
        return
      }
      setFile(selectedFile)
      setError(null)
      setDownloadUrl(null)
    }
  }

  const handleUpload = async () => {
    if (!file) {
      setError('Please select a file first')
      return
    }

    console.log('Starting upload for file:', file.name)
    setLoading(true)
    setError(null)
    setDownloadUrl(null)

    try {
      const formData = new FormData()
      formData.append('file', file)

      console.log('Sending request to /api/analyze')
      const response = await fetch('/api/analyze', {
        method: 'POST',
        body: formData,
      })

      console.log('Response received:', response.status, response.statusText)

      if (!response.ok) {
        let errorMessage = 'Failed to analyze document'
        try {
          const errorData = await response.json()
          errorMessage = errorData.error || errorMessage
        } catch (e) {
          // Response might not be JSON
          const text = await response.text()
          console.error('Error response:', text)
          errorMessage = text || errorMessage
        }
        throw new Error(errorMessage)
      }

      // Get the blob from response
      console.log('Creating download blob')
      const blob = await response.blob()
      console.log('Blob size:', blob.size, 'bytes')

      if (blob.size === 0) {
        throw new Error('Received empty document')
      }

      const url = window.URL.createObjectURL(blob)
      setDownloadUrl(url)
      console.log('Document ready for download')
    } catch (err) {
      console.error('Upload error:', err)
      setError(err instanceof Error ? err.message : 'An error occurred')
    } finally {
      setLoading(false)
    }
  }

  return (
    <main className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 py-12 px-4 sm:px-6 lg:px-8">
      <div className="max-w-3xl mx-auto">
        <div className="text-center mb-8">
          <h1 className="text-4xl font-bold text-gray-900 mb-2">
            Legal Document Analyzer
          </h1>
          <p className="text-lg text-gray-600">
            Upload a Word document (.docx) for AI-powered legal analysis
          </p>
        </div>

        <div className="bg-white shadow-xl rounded-lg p-8">
          <div className="space-y-6">
            {/* File Upload Section */}
            <div>
              <label
                htmlFor="file-upload"
                className="block text-sm font-medium text-gray-700 mb-2"
              >
                Select Document (.docx only)
              </label>
              <div className="mt-1 flex items-center space-x-4">
                <label className="flex-1 cursor-pointer">
                  <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 text-center hover:border-indigo-500 transition-colors">
                    <input
                      id="file-upload"
                      type="file"
                      accept=".docx"
                      onChange={handleFileChange}
                      className="hidden"
                      disabled={loading}
                    />
                    <div className="space-y-2">
                      <svg
                        className="mx-auto h-12 w-12 text-gray-400"
                        stroke="currentColor"
                        fill="none"
                        viewBox="0 0 48 48"
                        aria-hidden="true"
                      >
                        <path
                          d="M28 8H12a4 4 0 00-4 4v20m32-12v8m0 0v8a4 4 0 01-4 4H12a4 4 0 01-4-4v-4m32-4l-3.172-3.172a4 4 0 00-5.656 0L28 28M8 32l9.172-9.172a4 4 0 015.656 0L28 28m0 0l4 4m4-24h8m-4-4v8m-12 4h.02"
                          strokeWidth={2}
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        />
                      </svg>
                      <div className="text-sm text-gray-600">
                        {file ? (
                          <span className="font-medium text-indigo-600">
                            {file.name}
                          </span>
                        ) : (
                          <>
                            <span className="font-medium text-indigo-600 hover:text-indigo-500">
                              Click to upload
                            </span>
                            {' '}or drag and drop
                          </>
                        )}
                      </div>
                      <p className="text-xs text-gray-500">.docx files only</p>
                    </div>
                  </div>
                </label>
              </div>
            </div>

            {/* Error Message */}
            {error && (
              <div className="bg-red-50 border border-red-200 rounded-md p-4">
                <div className="flex">
                  <div className="flex-shrink-0">
                    <svg
                      className="h-5 w-5 text-red-400"
                      viewBox="0 0 20 20"
                      fill="currentColor"
                    >
                      <path
                        fillRule="evenodd"
                        d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z"
                        clipRule="evenodd"
                      />
                    </svg>
                  </div>
                  <div className="ml-3">
                    <p className="text-sm font-medium text-red-800">{error}</p>
                  </div>
                </div>
              </div>
            )}

            {/* Upload Button */}
            <button
              onClick={handleUpload}
              disabled={!file || loading}
              className={`w-full flex justify-center py-3 px-4 border border-transparent rounded-md shadow-sm text-sm font-medium text-white ${
                !file || loading
                  ? 'bg-gray-400 cursor-not-allowed'
                  : 'bg-indigo-600 hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500'
              } transition-colors`}
            >
              {loading ? (
                <div className="flex items-center">
                  <svg
                    className="animate-spin -ml-1 mr-3 h-5 w-5 text-white"
                    xmlns="http://www.w3.org/2000/svg"
                    fill="none"
                    viewBox="0 0 24 24"
                  >
                    <circle
                      className="opacity-25"
                      cx="12"
                      cy="12"
                      r="10"
                      stroke="currentColor"
                      strokeWidth="4"
                    />
                    <path
                      className="opacity-75"
                      fill="currentColor"
                      d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                    />
                  </svg>
                  Analyzing Document...
                </div>
              ) : (
                'Analyze Document'
              )}
            </button>

            {/* Download Section */}
            {downloadUrl && (
              <div className="bg-green-50 border border-green-200 rounded-md p-4">
                <div className="flex items-center justify-between">
                  <div className="flex items-center">
                    <svg
                      className="h-5 w-5 text-green-400"
                      viewBox="0 0 20 20"
                      fill="currentColor"
                    >
                      <path
                        fillRule="evenodd"
                        d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z"
                        clipRule="evenodd"
                      />
                    </svg>
                    <p className="ml-3 text-sm font-medium text-green-800">
                      Analysis complete! Document ready for download.
                    </p>
                  </div>
                  <a
                    href={downloadUrl}
                    download={`analyzed_${file?.name || 'document.docx'}`}
                    className="ml-4 inline-flex items-center px-4 py-2 border border-transparent text-sm font-medium rounded-md shadow-sm text-white bg-green-600 hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-green-500"
                  >
                    Download
                  </a>
                </div>
              </div>
            )}
          </div>
        </div>

        {/* Info Section */}
        <div className="mt-8 bg-white shadow rounded-lg p-6">
          <h2 className="text-lg font-semibold text-gray-900 mb-3">
            How it works:
          </h2>
          <ol className="list-decimal list-inside space-y-2 text-gray-600">
            <li>Upload your Word document (.docx format)</li>
            <li>AI analyzes the document for legal clarity and improvements</li>
            <li>Download the document with tracked changes and suggestions</li>
          </ol>
        </div>
      </div>
    </main>
  )
}
