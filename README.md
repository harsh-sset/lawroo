# Legal Document Analyzer

An AI-powered Next.js application that analyzes legal documents (.docx) and provides suggestions with track changes using Google Gemini.

## Features

- **File Upload**: Upload Word documents (.docx format only)
- **Original Formatting Preserved**: Maintains all fonts, styles, colors, and document structure
- **XML Manipulation**: Directly modifies the Word XML to add proper track changes
- **AI Analysis**: Uses Google Gemini to analyze legal documents for:
  - Legal clarity and precision
  - Ambiguous terms
  - Missing clauses
  - Potential legal risks
  - Grammar and formatting issues
- **Real Word Track Changes**: Adds actual Microsoft Word track changes (not visual replacements):
  - Uses `<w:del>` tags for deletions (shown with strikethrough in red)
  - Uses `<w:ins>` tags for insertions (shown with underline in blue)
  - Changes appear in Word's Review pane for accept/reject
  - Track changes are enabled automatically in the document
- **Inline Suggestions**: All changes are applied directly in the document at the exact location
- **Download**: Download the modified document with native Word track changes

## Prerequisites

- Node.js 18+ installed
- A Google AI API key (Gemini)

## Setup

1. **Clone or navigate to the project directory**

2. **Install dependencies**:
   ```bash
   npm install
   ```

3. **Configure environment variables**:
   - Copy `.env.example` to `.env.local`:
     ```bash
     cp .env.example .env.local
     ```
   - Add your Google Gemini API key to `.env.local`:
     ```
     GOOGLE_GENERATIVE_AI_API_KEY=your_actual_api_key_here
     ```
   - Get your API key from: https://aistudio.google.com/app/apikey

4. **Run the development server**:
   ```bash
   npm run dev
   ```

5. **Open your browser**:
   - Navigate to [http://localhost:3000](http://localhost:3000)

## Usage

1. Click the upload area or drag and drop a .docx file
2. Click "Analyze Document"
3. Wait for the AI analysis to complete (may take 10-30 seconds)
4. Download the document with native Word track changes

## Opening the Analyzed Document

When you open the downloaded document in Microsoft Word:
- Go to the **Review** tab
- Track changes will already be enabled
- You'll see deletions (strikethrough, red) and insertions (underline, blue)
- Use **Accept** or **Reject** buttons to review each change
- All original formatting (fonts, styles, headers, etc.) is preserved

## Tech Stack

- **Next.js 15**: React framework with App Router
- **TypeScript**: Type-safe development
- **Tailwind CSS**: Utility-first styling
- **Vercel AI SDK**: AI integration framework
- **Google Gemini**: AI model for legal analysis
- **mammoth**: Extract text from .docx files
- **jszip**: Parse and modify .docx file structure
- **xmldom**: Parse and manipulate WordprocessingML XML
- **xpath**: Query XML documents

## Project Structure

```
├── app/
│   ├── api/
│   │   └── analyze/
│   │       └── route.ts      # API endpoint for document processing
│   ├── globals.css           # Global styles with Tailwind
│   ├── layout.tsx            # Root layout
│   └── page.tsx              # Main page with file upload
├── .env.example              # Environment variables template
├── next.config.js            # Next.js configuration
├── package.json              # Dependencies
├── tailwind.config.js        # Tailwind configuration
└── tsconfig.json             # TypeScript configuration
```

## API Endpoint

### POST `/api/analyze`

Analyzes a legal document and returns a modified version with suggestions.

**Request**:
- Method: `POST`
- Content-Type: `multipart/form-data`
- Body: FormData with `file` field containing a .docx file

**Response**:
- Content-Type: `application/vnd.openxmlformats-officedocument.wordprocessingml.document`
- Body: Binary data of the modified .docx file

**Error Responses**:
- `400`: Invalid file or no file provided
- `500`: Server error or missing API key

## Environment Variables

- `GOOGLE_GENERATIVE_AI_API_KEY`: Your Google Gemini API key (required)

## Deployment

### Deploy to Vercel

1. Push your code to GitHub
2. Import the project in Vercel
3. Add the `GOOGLE_GENERATIVE_AI_API_KEY` environment variable in Vercel project settings
4. Deploy

### Other Platforms

This is a standard Next.js application and can be deployed to any platform that supports Next.js:
- Netlify
- AWS Amplify
- Railway
- Render
- Self-hosted with Node.js

## Limitations

- Only supports .docx files (not .doc or other formats)
- File size limited by Next.js server actions (default 10MB)
- API timeout set to 60 seconds
- Requires active internet connection for AI analysis

## License

ISC

## Support

For issues or questions, please open an issue in the project repository.
