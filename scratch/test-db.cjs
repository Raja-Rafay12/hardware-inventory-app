require('dotenv').config();
const { Pool } = require('pg');
const fs = require('fs');

const pool = new Pool({
  connectionString: process.env.PG_CONNECTION_STRING,
  ssl: { rejectUnauthorized: false }
});

const ALL_PRODUCTS = JSON.parse(fs.readFileSync('products.json', 'utf8'));

// Helper to remove malformed UTF-16 surrogate pairs that generate invalid UTF-8 byte sequences
function cleanString(val) {
  if (typeof val !== 'string') return val;
  try {
    let normalized = val.normalize('NFC');
    return normalized.replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|([^\uD800-\uDBFF]|^)[\uDC00-\uDFFF]/g, '$1');
  } catch (e) {
    return val;
  }
}

async function run() {
  const client = await pool.connect();
  try {
    console.log(`Loaded ${ALL_PRODUCTS.length} products. Testing bulk insertion...`);
    await client.query('BEGIN');
    
    const userId = '00000000-0000-0000-0000-000000000000'; // dummy uuid
    
    // Check if dummy user exists, if not create one
    await client.query(`
      INSERT INTO public.users (id, first_name, last_name, organization_name, email, password_hash)
      VALUES ($1, 'Test', 'User', 'Test Org', 'test@example.com', 'dummyhash')
      ON CONFLICT (email) DO NOTHING
    `, [userId]);

    const batchSize = 100;
    for (let i = 0; i < ALL_PRODUCTS.length; i += batchSize) {
      const batch = ALL_PRODUCTS.slice(i, i + batchSize);
      const valuePlaceholders = [];
      const queryParams = [userId];
      
      let paramIndex = 2;
      for (const p of batch) {
        valuePlaceholders.push(`($${paramIndex}, $1, $${paramIndex+1}, $${paramIndex+2}, $${paramIndex+3}, $${paramIndex+4}, $${paramIndex+5}, $${paramIndex+6}, $${paramIndex+7})`);
        queryParams.push(
          cleanString(p.id),
          cleanString(p.name || ''),
          cleanString(p.category || 'General'),
          cleanString(p.unit || 'piece'),
          Number(p.quantity || 0),
          Number(p.costPrice || p.cost_price || 0),
          Number(p.markup || p.markup || 40),
          Number(p.lowStock ?? p.low_stock ?? 5)
        );
        paramIndex += 8;
      }

      const queryText = `
        INSERT INTO public.products (id, user_id, name, category, unit, quantity, cost_price, markup, low_stock)
        VALUES ${valuePlaceholders.join(', ')}
        ON CONFLICT (user_id, id)
        DO UPDATE SET 
          name = EXCLUDED.name,
          category = EXCLUDED.category,
          unit = EXCLUDED.unit,
          quantity = EXCLUDED.quantity,
          cost_price = EXCLUDED.cost_price,
          markup = EXCLUDED.markup,
          low_stock = EXCLUDED.low_stock
      `;

      try {
        await client.query(queryText, queryParams);
      } catch (err) {
        console.error(`Error in batch starting at index ${i}:`);
        console.error(err);
        // Print the products in this batch
        console.error("Batch products sample:", batch.map(b => b.name).slice(0, 10));
        throw err;
      }
    }

    await client.query('COMMIT');
    console.log("Success! All 1849 products upserted without errors.");
  } catch (err) {
    console.error("General Failure:");
  } finally {
    client.release();
    await pool.end();
  }
}

run();
