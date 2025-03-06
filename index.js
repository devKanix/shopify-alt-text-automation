require('dotenv').config();
const XLSX = require('xlsx');
const { shopifyApi, ApiVersion } = require('@shopify/shopify-api');

// Initialize Shopify client
const shopify = shopifyApi({
  apiKey: process.env.SHOPIFY_ACCESS_TOKEN,
  apiVersion: ApiVersion.July23,
  isCustomStoreApp: true,
  adminApiAccessToken: process.env.SHOPIFY_ACCESS_TOKEN,
  hostName: `${process.env.SHOPIFY_SHOP_NAME}.myshopify.com`,
});

async function updateImageAltTexts() {
  try {
    // Read Excel file
    const workbook = XLSX.readFile('image-alt-texts.xlsx');
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(worksheet);

    console.log('Starting bulk update...');

    const client = new shopify.clients.Rest({
      session: {
        accessToken: process.env.SHOPIFY_ACCESS_TOKEN,
        shop: `${process.env.SHOPIFY_SHOP_NAME}.myshopify.com`,
      },
    });

    for (const row of data) {
      const { image_id, alt_text } = row;
      
      if (!image_id || !alt_text) {
        console.log(`Skipping row - missing required data: ${JSON.stringify(row)}`);
        continue;
      }

      try {
        // Update image alt text using Shopify API
        await client.put({
          path: `images/${image_id}`,
          data: {
            image: {
              id: image_id,
              alt: alt_text
            }
          }
        });

        console.log(`Updated image ${image_id} with alt text: ${alt_text}`);
      } catch (error) {
        console.error(`Error updating image ${image_id}:`, error.message);
      }

      // Add a small delay to respect API rate limits
      await new Promise(resolve => setTimeout(resolve, 500));
    }

    console.log('Bulk update completed!');
  } catch (error) {
    console.error('Error:', error.message);
  }
}

// Run the program
updateImageAltTexts();