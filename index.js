import { google } from 'googleapis';
import { v4 as uuidv4 } from 'uuid';
import fetch from 'node-fetch';
import dotenv from 'dotenv';
dotenv.config();

// Replace the hardcoded credentials with environment variables
const serviceAccount = {
  "type": "service_account",
  "project_id": process.env.GOOGLE_PROJECT_ID,
  "private_key_id": process.env.GOOGLE_PRIVATE_KEY_ID,
  "private_key": process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
  "client_email": process.env.GOOGLE_CLIENT_EMAIL,
  "client_id": process.env.GOOGLE_CLIENT_ID,
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": process.env.GOOGLE_CLIENT_X509_CERT_URL,
  "universe_domain": "googleapis.com"
};

const jwtClient = new google.auth.JWT(
  serviceAccount.client_email,
  null,
  serviceAccount.private_key,
  ['https://www.googleapis.com/auth/spreadsheets']
);

const sheets = google.sheets({ version: 'v4', auth: jwtClient });

// Define the sheet IDs for each day
const SHEET_IDS = {
  '2024-11-12': '1gWrSwjgfclJp0-VCW6GHjMOqbTfCn-Y0OGHYhZ5KpIU',
  '2024-11-13': '1shFpvIMJqEUMzeUcG8dL3EfdivDff5By1WgL_Ic-RqI',
  '2024-11-14': '1ag1-f51C7-40yBn5EDeamRnm5yw-ugIcvDvure32GEI',
  '2024-11-15': '1a58SeeQvXKfi_bTuymRcrmhm7nQ9w9K0gRqjp2Z6nMo',
  //'2024-11-12': '1v4mv1SfzOFeEYT1Au3IPzFE7dZ5FvRER6wMZA_lyF64'
};

const rooms = {
  'main-stage' : "MAINSTAGE / Masks",
  'stage-5' : "STAGE 5 / Hats",
  'stage-6' : "STAGE 6 / Kites",
  'stage-1' : "STAGE 1 / Fans",
  'stage-2' : "STAGE 2 / Lantern",
  'stage-3' : "STAGE 3 / Fabrics",
  'stage-4' : "STAGE 4 / Leafs",
  'classroom-a' : "CLASSROOM A",
  'classroom-b' : "CLASSROOM B",
  'classroom-c' : "CLASSROOM C",
  'classroom-d' : "CLASSROOM D",
  'classroom-e' : "CLASSROOM E",
  'breakout-1' : "BREAKOUT 1",
  'breakout-2' : "BREAKOUT 2",
  'breakout-3' : "BREAKOUT 3"

}

// Add this new function to fetch from the API
async function fetchScheduleData() {
  try {
    const response = await fetch('https://api.devcon.org/sessions?size=500&event=devcon-7', {
      headers: {
        'accept': 'application/json'
      }
    });
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    const result = await response.json();
    console.log(result.data.items.length);
    // Transform the API data to match the expected format
    return {
      rooms: Object.entries(rooms).map(([id, name]) => ({
        id,
        name: { en: name }
      })),
      talks: result.data.items.map(item => (
        {
        code: item.sourceId,
        title: { en: item.title },
        state: 'confirmed', // Assuming all API items are confirmed
        start: item.slot_start,
        end: item.slot_end,
        room: item.slot_roomId,
        speakers: item.speakers // Add speaker handling if available in API
      })),
    };
  } catch (error) {
    console.error('Error fetching schedule data:', error);
    throw error;
  }
}

async function main() {
  const scheduleData = await fetchScheduleData();
  
  // Loop through each day
  for (const [day, sheetId] of Object.entries(SHEET_IDS)) {
    console.log(`Processing sheet for ${day}`);
    
    // Update the SHEET_ID for the current day
    global.SHEET_ID = sheetId;
    
    try {
      // Create and populate overview sheet
      await createOrClearOverviewSheet(scheduleData.rooms);
      await populateOverviewSheet(scheduleData, scheduleData, day);
      
      // Create individual room sheets
      for (const [roomId, roomName] of Object.entries(rooms)) {
        try {
          await duplicateAndRenameSheet(roomName);  // Using roomName instead of roomId
          await populateSheetWithData(roomName, roomId, scheduleData, scheduleData, day);  // Pass both roomName and roomId
          console.log(`Completed processing room ${roomName} for ${day}`);
        } catch (error) {
          console.error(`Error processing room ${roomName} for ${day}:`, error);
          continue;
        }
      }
      
      console.log(`Completed processing for ${day}`);
    } catch (error) {
      console.error(`Error processing ${day}:`, error);
      continue;
    }
  }
}

// Create or clear the "overview" sheet
async function createOrClearOverviewSheet(rooms) {
  try {
    const sheetExists = await checkIfSheetExists('overview');
    if (!sheetExists) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: {
          requests: [{
            addSheet: {
              properties: {
                title: 'overview',
              },
            },
          }],
        },
      });
    } else {
      await sheets.spreadsheets.values.clear({
        spreadsheetId: SHEET_ID,
        range: 'overview',
      });
    }

    // Set the header row with room names
    const roomNames = rooms.map(room => room.name.en);
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: 'overview!A1',
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [roomNames],
      },
    });
  } catch (error) {
    console.error('Error in createOrClearOverviewSheet:', error);
    throw error;
  }
}

// Populate the "overview" sheet with session data
async function populateOverviewSheet(sessions, speakers, day) {
  const sessionTalks = sessions.talks.filter(talk => 
    talk.start && talk.start.split("T")[0] === day
  );
  const data = [];
  const mergeRequests = [];

  // Initialize headers
  const headers = ['Time', ...sessions.rooms.map(room => `${room.name.en}`)];
  data.push(headers);

  // Add a row for Room IC/Show Caller names
  const roomICRow = ['Room IC/Show Caller', ...Array(sessions.rooms.length).fill('')];
  data.push(roomICRow);

  // Initialize time slots
  for (let hour = 9; hour < 24; hour++) {
    for (let minute = 0; minute < 60; minute += 10) {
      const timeString = new Date(0, 0, 0, hour, minute).toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
      data.push([timeString, ...Array(sessions.rooms.length).fill('')]);
    }
  }

  // Get the sheet ID
  const overviewSheetId = await getSheetId('overview');
  if (!overviewSheetId) {
    console.error('Could not find overview sheet');
    return;
  }

  // Define formatting requests
  const formatRequests = [
    // Freeze first row and first column
    {
      updateSheetProperties: {
        properties: {
          sheetId: overviewSheetId,
          gridProperties: {
            frozenRowCount: 2, // Freeze both header and IC/Show Caller rows
            frozenColumnCount: 1
          }
        },
        fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount'
      }
    },
    // Set background color for header rows (both title and IC/Show Caller rows)
    {
      repeatCell: {
        range: {
          sheetId: overviewSheetId,
          startRowIndex: 0,
          endRowIndex: 2,
          startColumnIndex: 0,
          endColumnIndex: headers.length
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: {
              red: 0.9,
              green: 0.9,
              blue: 0.9
            },
            textFormat: {
              bold: true
            }
          }
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat)'
      }
    },
    // Set background color for first column
    {
      repeatCell: {
        range: {
          sheetId: overviewSheetId,
          startRowIndex: 0,
          endRowIndex: data.length,
          startColumnIndex: 0,
          endColumnIndex: 1
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: {
              red: 0.9,
              green: 0.9,
              blue: 0.9
            },
            textFormat: {
              bold: true
            }
          }
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat)'
      }
    },
    // Set column widths to 250px
    {
      updateDimensionProperties: {
        range: {
          sheetId: overviewSheetId,
          dimension: 'COLUMNS',
          startIndex: 0,
          endIndex: headers.length
        },
        properties: {
          pixelSize: 250
        },
        fields: 'pixelSize'
      }
    },
    // Set row height to 40px
    {
      updateDimensionProperties: {
        range: {
          sheetId: overviewSheetId,
          dimension: 'ROWS',
          startIndex: 0,
          endIndex: data.length
        },
        properties: {
          pixelSize: 40
        },
        fields: 'pixelSize'
      }
    },
    // Center align all cells
    {
      repeatCell: {
        range: {
          sheetId: overviewSheetId,
          startRowIndex: 0,
          endRowIndex: data.length,
          startColumnIndex: 0,
          endColumnIndex: headers.length
        },
        cell: {
          userEnteredFormat: {
            horizontalAlignment: 'CENTER',
            verticalAlignment: 'MIDDLE',
            wrapStrategy: 'WRAP'
          }
        },
        fields: 'userEnteredFormat(horizontalAlignment,verticalAlignment,wrapStrategy)'
      }
    }
  ];

  // First, unmerge all cells and apply formatting
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      requests: [
        {
          unmergeCells: {
            range: {
              sheetId: overviewSheetId,
              startRowIndex: 0,
              endRowIndex: data.length,
              startColumnIndex: 0,
              endColumnIndex: headers.length
            }
          }
        },
        ...formatRequests
      ]
    }
  });

  // Process sessions and create merge requests
  sessionTalks.forEach(session => {
    const roomIndex = sessions.rooms.findIndex(room => room.id === session.room) + 1;
    const startTime = new Date(session.start);
    const endTime = new Date(session.end);
    const durationMinutes = (endTime - startTime) / 60000;
    const startRow = Math.floor(((startTime.getHours() - 9) * 60 + startTime.getMinutes()) / 10) + 2;
    const numRows = Math.ceil(durationMinutes / 10);

    const speakerNames = getSpeakerNames(session.speakers, speakers.speakers);

    for (let i = 0; i < numRows; i++) {
      const row = startRow + i;
      if (row < data.length && roomIndex < data[row].length) {
        // Format with title and bold speakers
        data[row][roomIndex] = `${session.title.en ?? session.title}\n\n*${speakerNames}*`;
      }
    }

    if (numRows > 1) {
      mergeRequests.push({
        mergeCells: {
          range: {
            sheetId: overviewSheetId,
            startRowIndex: startRow,
            endRowIndex: startRow + numRows,
            startColumnIndex: roomIndex,
            endColumnIndex: roomIndex + 1
          },
          mergeType: 'MERGE_ALL'
        }
      });
    }
  });

  // Update the data
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: 'overview!A1',
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: data
    }
  });

  // Apply merge requests
  if (mergeRequests.length > 0) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: mergeRequests
      }
    });
  }
}

// Helper function to get speaker names
function getSpeakerNames(speakers) {
  return speakers?.map(speaker => speaker.name).join(', ') || '';
}

// Helper function to get the sheet ID by name
async function getSheetId(sheetName) {
  const sheet = await sheets.spreadsheets.get({
    spreadsheetId: SHEET_ID,
  });
  const sheetInfo = sheet.data.sheets.find((s) => s.properties.title === sheetName);
  return sheetInfo ? sheetInfo.properties.sheetId : null;
}

// Check if a sheet exists by name
async function checkIfSheetExists(sheetName) {
  const sheet = await sheets.spreadsheets.get({
    spreadsheetId: SHEET_ID,
  });
  return sheet.data.sheets.some((s) => s.properties.title === sheetName);
}

// Function to duplicate a sheet and rename it
async function duplicateAndRenameSheet(newSheetName) {
  try {
    // Check if sheet already exists
    const sheetExists = await checkIfSheetExists(newSheetName);
    if (sheetExists) {
      console.log(`Sheet ${newSheetName} already exists, skipping creation`);
      return;
    }

    // Duplicate the template sheet
    const copyResponse = await sheets.spreadsheets.sheets.copyTo({
      spreadsheetId: SHEET_ID,
      sheetId: 688800800, // ID of the template sheet
      requestBody: {
        destinationSpreadsheetId: SHEET_ID,
      },
    });

    const newSheetId = copyResponse.data.sheetId;

    // Rename the duplicated sheet and apply formatting
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: [
          {
            updateSheetProperties: {
              properties: {
                sheetId: newSheetId,
                title: newSheetName,
              },
              fields: 'title',
            },
          },
          // Set row height to 30px
          {
            updateDimensionProperties: {
              range: {
                sheetId: newSheetId,
                dimension: 'ROWS',
                startIndex: 0,
              },
              properties: {
                pixelSize: 30
              },
              fields: 'pixelSize'
            }
          },
          // Add black borders only to cells starting from row 8
          {
            updateBorders: {
              range: {
                sheetId: newSheetId,
                startRowIndex: 7,  // 0-based index, so 7 is row 8
                startColumnIndex: 0,
                endColumnIndex: 20  // Assuming 20 columns (A-T)
              },
              top: {
                style: 'SOLID',
                color: { red: 0, green: 0, blue: 0 }
              },
              bottom: {
                style: 'SOLID',
                color: { red: 0, green: 0, blue: 0 }
              },
              left: {
                style: 'SOLID',
                color: { red: 0, green: 0, blue: 0 }
              },
              right: {
                style: 'SOLID',
                color: { red: 0, green: 0, blue: 0 }
              },
              innerHorizontal: {
                style: 'SOLID',
                color: { red: 0, green: 0, blue: 0 }
              },
              innerVertical: {
                style: 'SOLID',
                color: { red: 0, green: 0, blue: 0 }
              }
            }
          }
        ],
      },
    });

    console.log(`Sheet duplicated and renamed to: ${newSheetName}`);
  } catch (error) {
    console.error('Error duplicating and renaming sheet:', error.message);
    throw error;
  }
}


const merge = (a, b, predicate = (a, b) => a === b) => {
  if (!a) return b;
  
  const c = [...a]; // copy to avoid side effects
  
  // Ensure arrays don't exceed 20 elements (columns A-T)
  c.length = Math.min(c.length, 20);
  b.length = Math.min(b.length, 20);

  // Replace the row ID (second element) from b
  if (b.length > 1) {
    c[1] = b[1];
  }

  // Determine the limit for the loop to exclude the last 8 items
  const limit = Math.max(b.length - 8, 0);

  // Add all items from B to copy C if they're not already present, except the row ID
  b.slice(0, limit).forEach((bItem, index) => {
    if (index !== 1 && !c.some((cItem) => predicate(bItem, cItem))) {
      c.push(bItem);
    }
  });

  // Handle the last 8 positions
  const lastPositionsStart = Math.max(c.length - 8, 0);
  for (let i = lastPositionsStart; i < Math.min(c.length, 20); i++) {
    c[i] = b[i] !== undefined ? b[i] : ""; // Preserve b's value or replace with empty string
  }

  // Final length check to ensure we don't exceed 20 columns
  return c.slice(0, 20);
}


// Populate the sheet with data for the given room and day
async function populateSheetWithData(sheetName, roomId, sessions, speakers, day) {
  const sessionTalks = sessions.talks.filter(talk => 
    talk.room === roomId && 
    talk.start.split('T')[0] === day
  ).sort((a, b) => new Date(a.start) - new Date(b.start));

  if (sessionTalks.length === 0) {
    console.log(`No sessions found for room ${sheetName} on ${day}`);
    return;
  }

  const currentData = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `'${sheetName}'!A1:T${sessionTalks.length + 8}`,
  });

  const data = sessionTalks.map((session, index) => {
    // Find the existing row with matching session code
    const currentRow = currentData.data.values?.find(row => row[0] === session.code);
    const newRow = [
      session.code,
      index + 1, // Always use new index
      new Date(session.start).toLocaleTimeString('en-US', { hour: 'numeric', minute: 'numeric', hour12: true }),
      `${(new Date(session.end).getTime() - new Date(session.start).getTime()) / 60000} minutes`,
      new Date(session.end).toLocaleTimeString('en-US', { hour: 'numeric', minute: 'numeric', hour12: true }),
      session.title.en,
      getSpeakerNames(session.speakers),
      `https://devcon.org/sea/presentation/${session.code}`,
      '-',
      '-',
      '-',
      '-',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      ''
    ];
    console.log(currentRow ? merge(newRow, currentRow) : newRow);
    return {
      range: `'${sheetName}'!A${index + 8}:T${index + 8}`,
      values: [currentRow ? merge(newRow, currentRow) : newRow]
    };
  });

  try {
    await bulkUpdateData(data);
    
    // Get the sheet ID for formatting
    const sheet = await sheets.spreadsheets.get({
      spreadsheetId: SHEET_ID,
    });
    const sheetInfo = sheet.data.sheets.find((s) => s.properties.title === sheetName);
    const sheetId = sheetInfo.properties.sheetId;

    // Apply formatting after data is populated
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: [
          // Set row height to 50px starting from row 5
          {
            updateDimensionProperties: {
              range: {
                sheetId: sheetId,
                dimension: 'ROWS',
                startIndex: 4, // Row 5 (0-indexed)
                endIndex: sessionTalks.length + 8 // Include header rows
              },
              properties: {
                pixelSize: 50
              },
              fields: 'pixelSize'
            }
          },
          // Add black borders to all cells starting from row 5
          {
            updateBorders: {
              range: {
                sheetId: sheetId,
                startRowIndex: 4, // Row 5 (0-indexed)
                endRowIndex: sessionTalks.length + 8, // Include header rows
                startColumnIndex: 0,
                endColumnIndex: 20  // Columns A-T
              },
              top: {
                style: 'SOLID',
                color: { red: 0, green: 0, blue: 0 }
              },
              bottom: {
                style: 'SOLID',
                color: { red: 0, green: 0, blue: 0 }
              },
              left: {
                style: 'SOLID',
                color: { red: 0, green: 0, blue: 0 }
              },
              right: {
                style: 'SOLID',
                color: { red: 0, green: 0, blue: 0 }
              },
              innerHorizontal: {
                style: 'SOLID',
                color: { red: 0, green: 0, blue: 0 }
              },
              innerVertical: {
                style: 'SOLID',
                color: { red: 0, green: 0, blue: 0 }
              }
            }
          }
        ]
      }
    });

    console.log(`Successfully populated and formatted data for ${sheetName}`);
  } catch (error) {
    console.error(`Error populating data for ${sheetName}:`, error);
    throw error;
  }
}

// Bulk update data in the Google Sheet
async function bulkUpdateData(data) {
  if (data.length === 0) return;

  try {
    const response = await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SHEET_ID,
      valueInputOption: 'USER_ENTERED',
      requestBody: { data },
    });
    console.log('Data successfully updated:');
  } catch (error) {
    console.error('Error updating data:', error.message);
  }
}

// Helper function to check if a range is already merged
function isRangeAlreadyMerged(range, mergeRequests) {
  return mergeRequests.some(request => {
    const existingRange = request.mergeCells.range;
    return (
      existingRange.startRowIndex === range.startRowIndex &&
      existingRange.endRowIndex === range.endRowIndex &&
      existingRange.startColumnIndex === range.startColumnIndex &&
      existingRange.endColumnIndex === range.endColumnIndex
    );
  });
}

// Start the process
main();

