require('dotenv').config();
const { Pool } = require('pg');

const pool = new Pool({
  connectionString: process.env.PG_CONNECTION_STRING,
  ssl: { rejectUnauthorized: false }
});

const SEED_PRODUCTS = [
  { id: "p0001", name: ".", category: "General", unit: "piece", quantity: 0, costPrice: 380.0, markup: 40.0, lowStock: 5 },
  { id: "p0073", name: "BAILCHA  4#", category: "Hardware & Tools", unit: "piece", quantity: 0, costPrice: 400.0, markup: 40.0, lowStock: 5 }
];

async function run() {
  const client = await pool.connect();
  try {
    console.log("Connected to cloud database. Testing products upsert...");
    await client.query('BEGIN');
    
    // We will simulate the batch query for these two products
    const userId = '00000000-0000-0000-0000-000000000000'; // dummy uuid
    
    // Check if dummy user exists, if not create one
    await client.query(`
      INSERT INTO public.users (id, first_name, last_name, organization_name, email, password_hash)
      VALUES ($1, 'Test', 'User', 'Test Org', 'test@example.com', 'dummyhash')
      ON CONFLICT (email) DO NOTHING
    `, [userId]);

    const batch = SEED_PRODUCTS;
    const valuePlaceholders = [];
    const queryParams = [userId];
    
    let paramIndex = 2;
    for (const p of batch) {
      valuePlaceholders.push(`($${paramIndex}, $1, $${paramIndex+1}, $${paramIndex+2}, $${paramIndex+3}, $${paramIndex+4}, $${paramIndex+5}, $${paramIndex+6}, $${paramIndex+7})`);
      queryParams.push(
        p.id,
        p.name || '',
        p.category || 'General',
        p.unit || 'piece',
        Number(p.quantity || 0),
        Number(p.costPrice || 0),
        Number(p.markup || 40),
        Number(p.lowStock ?? 5)
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

    console.log("Running SQL Query...");
    await client.query(queryText, queryParams);
    await client.query('COMMIT');
    console.log("Success! Products upsert completed without errors.");
  } catch (err) {
    console.error("SQL Error encountered:");
    console.error(err);
  } finally {
    client.release();
    await pool.end();
  }
}

run();
