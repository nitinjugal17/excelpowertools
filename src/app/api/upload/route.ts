
import { NextRequest, NextResponse } from 'next/server';
import path from 'path';
import fs from 'fs/promises';

export async function POST(request: NextRequest) {
  if (process.env.SAVE_UPLOADS_TO_SERVER !== 'true') {
    return NextResponse.json({ message: 'File saving is disabled.' }, { status: 200 });
  }

  try {
    const formData = await request.formData();
    // Handle both single ('file') and multiple ('files[]') file uploads gracefully
    const files = formData.getAll('file').concat(formData.getAll('files[]'));

    if (files.length === 0 || !(files[0] instanceof File)) {
      return NextResponse.json({ error: 'No files were uploaded.' }, { status: 400 });
    }

    // Use a local 'uploads' directory. The next.config.js is configured to ignore this.
    const uploadDir = path.join(process.cwd(), 'uploads');
    
    // Ensure the directory exists
    await fs.mkdir(uploadDir, { recursive: true });
    
    for (const file of files) {
      if (file instanceof File) {
        const buffer = Buffer.from(await file.arrayBuffer());
        // Sanitize file name to prevent directory traversal attacks
        const sanitizedFilename = path.basename(file.name).replace(/\\|\//g, '');
        if (!sanitizedFilename) continue;

        const filePath = path.join(uploadDir, sanitizedFilename);
        await fs.writeFile(filePath, buffer);
      }
    }

    return NextResponse.json({ message: 'Files uploaded successfully.' }, { status: 200 });
  } catch (error) {
    console.error('Error saving uploaded file:', error);
    return NextResponse.json({ error: 'Failed to save file.' }, { status: 500 });
  }
}
