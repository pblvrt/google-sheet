const axios = require('axios');

async function extractSlotRooms() {
  try {
    // Make the API call
    const response = await axios.get('https://api.devcon.org/sessions?event=devcon-7', {
      headers: {
        'accept': 'application/json'
      }
    });

    console.log(response.data);

    // Extract all slot_roomId combinations
    const slotRooms = response.data.data.items
      .map(item => item.slot_roomId)
      .filter(Boolean) // Remove any null/undefined values
      .sort(); // Sort alphabetically

    // Remove duplicates using Set
    const uniqueSlotRooms = [...new Set(slotRooms)];

    // Print results
    console.log('All unique slot_roomId combinations:');
    uniqueSlotRooms.forEach(id => console.log(id));
    console.log(`\nTotal unique combinations: ${uniqueSlotRooms.length}`);

    // Optional: Save to a file
    const fs = require('fs');
    fs.writeFileSync('slot-rooms.json', JSON.stringify(uniqueSlotRooms, null, 2));
    console.log('\nResults have been saved to slot-rooms.json');

  } catch (error) {
    console.error('Error fetching or processing data:', error.message);
  }
}

extractSlotRooms();