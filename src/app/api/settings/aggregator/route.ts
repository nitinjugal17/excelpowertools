
import { NextRequest, NextResponse } from 'next/server';
import path from 'path';
import fs from 'fs/promises';

// Define the path to the settings file
const settingsFilePath = path.join(process.cwd(), 'src', 'data', 'aggregator-settings.json');
const dataDir = path.dirname(settingsFilePath);

/**
 * Handles GET requests to retrieve the aggregator settings.
 */
export async function GET() {
  try {
    const data = await fs.readFile(settingsFilePath, 'utf-8');
    return NextResponse.json(JSON.parse(data));
  } catch (error: any) {
    if (error.code === 'ENOENT') {
      // If the file doesn't exist, return empty settings, client will use defaults.
      return NextResponse.json({});
    }
    console.error('Failed to read settings:', error);
    return NextResponse.json({ error: 'Failed to read settings file.' }, { status: 500 });
  }
}

/**
 * Handles POST requests to save the aggregator settings.
 */
export async function POST(request: NextRequest) {
  try {
    const settings = await request.json();
    
    // Ensure the data directory exists
    await fs.mkdir(dataDir, { recursive: true });
    
    // Write the new settings to the file
    await fs.writeFile(settingsFilePath, JSON.stringify(settings, null, 2));
    
    return NextResponse.json({ message: 'Settings saved successfully.' });
  } catch (error) {
    console.error('Failed to save settings:', error);
    return NextResponse.json({ error: 'Failed to save settings file.' }, { status: 500 });
  }
}
