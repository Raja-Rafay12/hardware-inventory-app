const fs = require('fs');
const path = require('path');

// Try loading from next to the executable first, then fall back to the packaged .env
const externalEnvPath = path.join(path.dirname(process.execPath), '.env');
if (fs.existsSync(externalEnvPath)) {
  require('dotenv').config({ path: externalEnvPath });
} else {
  require('dotenv').config({ path: path.join(__dirname, '.env') });
}

const { app, BrowserWindow, ipcMain } = require('electron');
const { Pool } = require('pg');
const bcrypt = require('bcryptjs');
const nodemailer = require('nodemailer');

// Helper function to send confirmation/welcome email via Google SMTP protocol
async function sendWelcomeEmail(toEmail, firstName, organizationName) {
  const host = process.env.SMTP_HOST;
  const port = process.env.SMTP_PORT;
  const secure = process.env.SMTP_SECURE === 'true';
  const user = process.env.SMTP_USER;
  const pass = process.env.SMTP_PASS;
  const from = process.env.SMTP_FROM || `"Hardware Inventory" <${user}>`;

  if (!host || !user || !pass) {
    console.log("Google SMTP not configured in .env. Skipping welcome email dispatch.");
    return;
  }

  try {
    const transporter = nodemailer.createTransport({
      host,
      port: parseInt(port) || 465,
      secure,
      auth: {
        user,
        pass
      }
    });

    const htmlContent = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 25px; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; color: #333333;">
        <div style="text-align: center; margin-bottom: 20px;">
          <h2 style="color: #d9720b; margin: 0; padding-bottom: 10px; border-bottom: 2px solid #d9720b;">Registration Confirmed</h2>
        </div>
        <p>Hi <strong>${firstName}</strong>,</p>
        <p>Thank you for signing up for the Hardware Inventory App! Your account has been successfully created, and your organization profile for <strong>${organizationName}</strong> is ready.</p>
        <p>You can now manage your catalog, track stock, and log invoice transactions seamlessly in the cloud.</p>
        <br/>
        <div style="background-color: #fcfbf7; border-left: 4px solid #d9720b; padding: 15px; font-size: 13px; color: #666666; border-radius: 4px; line-height: 1.5;">
          <strong>Account Security Information:</strong> This email confirms your credentials have been verified and your profile is locked to your private secure workspace. If you did not register for this service, please contact your administrator.
        </div>
        <br/>
        <p style="font-size: 12px; color: #999999; text-align: center; border-top: 1px solid #eeeeee; padding-top: 15px; margin-bottom: 0;">
          Sent automatically via Google SMTP.
        </p>
      </div>
    `;

    await transporter.sendMail({
      from,
      to: toEmail,
      subject: "Welcome to Hardware Inventory App - Account Confirmed",
      html: htmlContent
    });
    console.log(`Welcome email successfully sent via Google SMTP to ${toEmail}`);
  } catch (err) {
    console.error("Failed to send welcome email via Google SMTP:", err);
  }
}

// Map to cache OTPs: Email -> { otpCode, expiresAt }
const resetOtps = new Map();

// Helper to send password reset verification email via Google SMTP
async function sendResetPasswordEmail(toEmail, otpCode) {
  const host = process.env.SMTP_HOST;
  const port = process.env.SMTP_PORT;
  const secure = process.env.SMTP_SECURE === 'true';
  const user = process.env.SMTP_USER;
  const pass = process.env.SMTP_PASS;
  const from = process.env.SMTP_FROM || `"Hardware Inventory" <${user}>`;

  if (!host || !user || !pass) {
    console.log("Google SMTP not configured in .env. Skipping password reset email dispatch.");
    throw new Error("SMTP server is not configured. Cannot send reset email.");
  }

  try {
    const transporter = nodemailer.createTransport({
      host,
      port: parseInt(port) || 465,
      secure,
      auth: {
        user,
        pass
      }
    });

    const htmlContent = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 25px; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; color: #333333;">
        <div style="text-align: center; margin-bottom: 20px;">
          <h2 style="color: #d9720b; margin: 0; padding-bottom: 10px; border-bottom: 2px solid #d9720b;">Password Reset Verification</h2>
        </div>
        <p>Hello,</p>
        <p>We received a request to reset the password for your account associated with this email address.</p>
        <div style="text-align: center; margin: 30px 0;">
          <span style="font-size: 28px; font-weight: 700; color: #d9720b; letter-spacing: 4px; background-color: #fcfbf7; border: 1px dashed #d9720b; padding: 10px 20px; border-radius: 4px;">${otpCode}</span>
        </div>
        <p>This verification code is valid for <strong>10 minutes</strong>. If you did not make this request, you can safely ignore this email; your password will remain unchanged.</p>
        <br/>
        <div style="background-color: #fcfbf7; border-left: 4px solid #d9720b; padding: 15px; font-size: 13px; color: #666666; border-radius: 4px; line-height: 1.5;">
          <strong>Security Notice:</strong> Never share this verification code with anyone. App administrators will never ask for this code.
        </div>
        <br/>
        <p style="font-size: 12px; color: #999999; text-align: center; border-top: 1px solid #eeeeee; padding-top: 15px; margin-bottom: 0;">
          Sent automatically via Google SMTP.
        </p>
      </div>
    `;

    await transporter.sendMail({
      from,
      to: toEmail,
      subject: "Hardware Inventory App - Password Reset Verification Code",
      html: htmlContent
    });
    console.log(`Password reset verification email successfully sent to ${toEmail}`);
  } catch (err) {
    console.error("Failed to send password reset email via Google SMTP:", err);
    throw err;
  }
}


// Helper to remove null bytes and malformed UTF-16 surrogate pairs that generate invalid UTF-8 byte sequences
function cleanString(val) {
  if (typeof val !== 'string') return val;
  try {
    let cleaned = val.replace(/\u0000/g, '');
    let normalized = cleaned.normalize('NFC');
    return normalized.replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|([^\uD800-\uDBFF]|^)[\uDC00-\uDFFF]/g, '$1');
  } catch (e) {
    return val;
  }
}

// Initialize database pool
const connectionString = process.env.PG_CONNECTION_STRING;
if (!connectionString) {
  console.error("FATAL ERROR: PG_CONNECTION_STRING is not defined in .env file.");
}

const pool = new Pool({
  connectionString: connectionString,
  ssl: {
    rejectUnauthorized: false // Required for secure cloud PostgreSQL hosts (like Neon)
  }
});

let currentUserId = null;

function createWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 800,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.cjs')
    },
    autoHideMenuBar: true,
  });

  // Open Developer Tools for debugging console errors
  // win.webContents.openDevTools();

  // Load the built index.html
  win.loadFile(path.join(__dirname, 'dist', 'index.html'));
}

// ==========================================
// ELECTRON IPC AUTHENTICATION HANDLERS
// ==========================================

ipcMain.handle('auth:signup', async (event, signupData) => {
  const client = await pool.connect();
  try {
    const { firstName, lastName, organizationName, email, password } = signupData;
    const formattedEmail = String(email).trim().toLowerCase();
    
    // Check if user already exists
    const checkUser = await client.query('SELECT id FROM public.users WHERE email = $1', [formattedEmail]);
    if (checkUser.rows.length > 0) {
      throw new Error('An account with this email address already exists.');
    }

    // Hash password
    const passwordHash = await bcrypt.hash(password, 10);

    await client.query('BEGIN');

    // Insert user
    const insertResult = await client.query(
      `INSERT INTO public.users(first_name, last_name, organization_name, email, password_hash)
       VALUES ($1, $2, $3, $4, $5)
       RETURNING id, email, first_name, last_name, organization_name`,
      [firstName.trim(), lastName.trim(), organizationName.trim(), formattedEmail, passwordHash]
    );

    const user = insertResult.rows[0];
    currentUserId = user.id;

    // Seed default settings automatically using their organization name as the shop name
    await client.query(
      `INSERT INTO public.settings (user_id, shop_name, currency_symbol, invoice_counter, low_stock_default)
       VALUES ($1, $2, $3, $4, $5)`,
      [currentUserId, user.organization_name, 'Rs ', 1, 5]
    );

    await client.query('COMMIT');

    // Send the email in the background (non-blocking)
    sendWelcomeEmail(user.email, user.first_name, user.organization_name).catch(console.error);

    return {
      id: user.id,
      email: user.email,
      firstName: user.first_name,
      lastName: user.last_name,
      organizationName: user.organization_name
    };
  } catch (err) {
    await client.query('ROLLBACK');
    console.error("Signup error:", err);
    throw new Error(err.message || 'Failed to create account.');
  } finally {
    client.release();
  }
});

ipcMain.handle('auth:login', async (event, email, password) => {
  try {
    const formattedEmail = String(email).trim().toLowerCase();

    // Query user
    const result = await pool.query('SELECT * FROM public.users WHERE email = $1', [formattedEmail]);
    if (result.rows.length === 0) {
      throw new Error('Invalid email or password.');
    }

    const user = result.rows[0];
    
    // Verify password hash
    const match = await bcrypt.compare(password, user.password_hash);
    if (!match) {
      throw new Error('Invalid email or password.');
    }

    currentUserId = user.id;
    return {
      id: user.id,
      email: user.email,
      firstName: user.first_name,
      lastName: user.last_name,
      organizationName: user.organization_name
    };
  } catch (err) {
    console.error("Login error:", err);
    throw new Error(err.message || 'Authentication failed.');
  }
});

ipcMain.handle('auth:logout', () => {
  currentUserId = null;
  return { success: true };
});

ipcMain.handle('auth:request-reset-otp', async (event, email) => {
  try {
    const formattedEmail = String(email).trim().toLowerCase();
    
    // Check if the user exists
    const result = await pool.query('SELECT id FROM public.users WHERE email = $1', [formattedEmail]);
    if (result.rows.length === 0) {
      throw new Error('No account found with this email address.');
    }
    
    // Generate a 6-digit random code
    const otpCode = Math.floor(100000 + Math.random() * 900000).toString();
    const expiresAt = Date.now() + 10 * 60 * 1000; // 10 minutes expiry
    
    // Store in cache
    resetOtps.set(formattedEmail, { otpCode, expiresAt });
    
    // Send email
    await sendResetPasswordEmail(formattedEmail, otpCode);
    
    return { success: true, message: 'Verification email sent. Please check your inbox.' };
  } catch (err) {
    console.error("Request reset OTP error:", err);
    throw new Error(err.message || 'Failed to send verification code.');
  }
});

ipcMain.handle('auth:reset-password-confirm', async (event, email, otp, newPassword) => {
  try {
    const formattedEmail = String(email).trim().toLowerCase();
    const cleanOtp = String(otp).trim();
    
    // Check cache
    const cacheEntry = resetOtps.get(formattedEmail);
    if (!cacheEntry) {
      throw new Error('No reset request found for this email address. Please request a new code.');
    }
    
    // Check expiry
    if (Date.now() > cacheEntry.expiresAt) {
      resetOtps.delete(formattedEmail);
      throw new Error('Verification code has expired. Please request a new code.');
    }
    
    // Match OTP code
    if (cacheEntry.otpCode !== cleanOtp) {
      throw new Error('Invalid verification code.');
    }
    
    // Hash password
    const passwordHash = await bcrypt.hash(newPassword, 10);
    
    // Update user in database
    await pool.query('UPDATE public.users SET password_hash = $1 WHERE email = $2', [passwordHash, formattedEmail]);
    
    // Clear cache entry
    resetOtps.delete(formattedEmail);
    
    return { success: true, message: 'Password has been reset successfully.' };
  } catch (err) {
    console.error("Confirm reset password error:", err);
    throw new Error(err.message || 'Failed to reset password.');
  }
});


ipcMain.handle('auth:current-user', async () => {
  if (!currentUserId) return null;
  try {
    const result = await pool.query(
      'SELECT id, email, first_name as "firstName", last_name as "lastName", organization_name as "organizationName" FROM public.users WHERE id = $1',
      [currentUserId]
    );
    return result.rows[0] || null;
  } catch (err) {
    console.error("GetCurrentUser error:", err);
    return null;
  }
});

// ==========================================
// ELECTRON IPC DATABASE CRUD HANDLERS
// ==========================================

ipcMain.handle('db:fetch-data', async () => {
  if (!currentUserId) throw new Error('Unauthorized');
  try {
    // 1. Fetch products
    const productsResult = await pool.query(
      'SELECT * FROM public.products WHERE user_id = $1 ORDER BY name ASC',
      [currentUserId]
    );

    // 2. Fetch invoices
    const invoicesResult = await pool.query(
      'SELECT * FROM public.invoices WHERE user_id = $1 ORDER BY date DESC',
      [currentUserId]
    );

    // 3. Fetch expenses
    const expensesResult = await pool.query(
      'SELECT * FROM public.expenses WHERE user_id = $1 ORDER BY date DESC',
      [currentUserId]
    );

    // 4. Fetch settings
    const settingsResult = await pool.query(
      'SELECT * FROM public.settings WHERE user_id = $1',
      [currentUserId]
    );

    // Format DB objects to match JS expectations
    const products = productsResult.rows.map(p => ({
      id: p.id,
      name: p.name,
      category: p.category,
      unit: p.unit,
      quantity: Number(p.quantity),
      costPrice: Number(p.cost_price),
      markup: Number(p.markup),
      lowStock: p.low_stock
    }));

    const invoices = invoicesResult.rows.map(i => ({
      id: i.id,
      invoiceNumber: i.invoice_number,
      customerName: i.customer_name,
      date: i.date.toISOString(),
      items: i.items,
      subtotal: Number(i.subtotal),
      discount: Number(i.discount),
      total: Number(i.total),
      totalCost: Number(i.total_cost),
      profit: Number(i.profit)
    }));

    const expenses = expensesResult.rows.map(e => ({
      id: e.id,
      date: e.date.toISOString().split('T')[0], // yyyy-mm-dd
      description: e.description,
      category: e.category,
      amount: Number(e.amount)
    }));

    const dbSettings = settingsResult.rows[0];
    const settings = dbSettings ? {
      shopName: dbSettings.shop_name,
      currencySymbol: dbSettings.currency_symbol,
      invoiceCounter: dbSettings.invoice_counter,
      lowStockDefault: dbSettings.low_stock_default
    } : null;

    return { products, invoices, expenses, settings };
  } catch (err) {
    console.error("Fetch shop data error:", err);
    throw new Error('Failed to load shop data from cloud database.');
  }
});

ipcMain.handle('db:save-products', async (event, products) => {
  if (!currentUserId) throw new Error('Unauthorized');
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    await client.query("SET client_encoding TO 'UTF8'");

    // Find deleted products by comparing incoming array with existing rows
    const currentRows = await client.query('SELECT id FROM public.products WHERE user_id = $1', [currentUserId]);
    const currentDbIds = currentRows.rows.map(r => r.id);
    const incomingIds = products.map(p => p.id);
    const deletedIds = currentDbIds.filter(id => !incomingIds.includes(id));

    if (deletedIds.length > 0) {
      await client.query(
        'DELETE FROM public.products WHERE user_id = $1 AND id = ANY($2::text[])',
        [currentUserId, deletedIds]
      );
    }

    // Upsert products in optimized batches of 100 to reduce latency round-trips
    const batchSize = 100;
    for (let i = 0; i < products.length; i += batchSize) {
      const batch = products.slice(i, i + batchSize);
      const valuePlaceholders = [];
      const queryParams = [currentUserId];
      
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
      
      await client.query(queryText, queryParams);
    }

    await client.query('COMMIT');
    return { success: true };
  } catch (err) {
    await client.query('ROLLBACK');
    console.error("Save products error:", err);
    throw new Error('Failed to save products to database.');
  } finally {
    client.release();
  }
});

ipcMain.handle('db:save-invoices', async (event, invoices) => {
  if (!currentUserId) throw new Error('Unauthorized');
  const client = await pool.connect();
  try {
    await client.query('BEGIN');

    // Find deleted invoices
    const currentRows = await client.query('SELECT id FROM public.invoices WHERE user_id = $1', [currentUserId]);
    const currentDbIds = currentRows.rows.map(r => r.id);
    const incomingIds = invoices.map(i => i.id);
    const deletedIds = currentDbIds.filter(id => !incomingIds.includes(id));

    if (deletedIds.length > 0) {
      await client.query(
        'DELETE FROM public.invoices WHERE user_id = $1 AND id = ANY($2::text[])',
        [currentUserId, deletedIds]
      );
    }

    // Insert new invoices (invoices are read-only, so only need to insert new ones)
    const newInvoices = invoices.filter(i => !currentDbIds.includes(i.id));
    for (const inv of newInvoices) {
      await client.query(
        `INSERT INTO public.invoices (id, user_id, invoice_number, customer_name, date, subtotal, discount, total, total_cost, profit, items)
         VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11)`,
        [inv.id, currentUserId, inv.invoiceNumber, inv.customerName, inv.date, inv.subtotal, inv.discount, inv.total, inv.totalCost, inv.profit, JSON.stringify(inv.items)]
      );
    }

    await client.query('COMMIT');
    return { success: true };
  } catch (err) {
    await client.query('ROLLBACK');
    console.error("Save invoices error:", err);
    throw new Error('Failed to save invoices to database.');
  } finally {
    client.release();
  }
});

ipcMain.handle('db:save-expenses', async (event, expenses) => {
  if (!currentUserId) throw new Error('Unauthorized');
  const client = await pool.connect();
  try {
    await client.query('BEGIN');

    // Find deleted expenses
    const currentRows = await client.query('SELECT id FROM public.expenses WHERE user_id = $1', [currentUserId]);
    const currentDbIds = currentRows.rows.map(r => r.id);
    const incomingIds = expenses.map(e => e.id);
    const deletedIds = currentDbIds.filter(id => !incomingIds.includes(id));

    if (deletedIds.length > 0) {
      await client.query(
        'DELETE FROM public.expenses WHERE user_id = $1 AND id = ANY($2::text[])',
        [currentUserId, deletedIds]
      );
    }

    // Insert new expenses
    const newExpenses = expenses.filter(e => !currentDbIds.includes(e.id));
    for (const exp of newExpenses) {
      await client.query(
        `INSERT INTO public.expenses (id, user_id, date, description, category, amount)
         VALUES ($1, $2, $3, $4, $5, $6)`,
        [exp.id, currentUserId, exp.date, exp.description, exp.category, exp.amount]
      );
    }

    await client.query('COMMIT');
    return { success: true };
  } catch (err) {
    await client.query('ROLLBACK');
    console.error("Save expenses error:", err);
    throw new Error('Failed to save expenses to database.');
  } finally {
    client.release();
  }
});

ipcMain.handle('db:save-settings', async (event, settings) => {
  if (!currentUserId) throw new Error('Unauthorized');
  try {
    await pool.query(
      `INSERT INTO public.settings (user_id, shop_name, currency_symbol, invoice_counter, low_stock_default)
       VALUES ($1, $2, $3, $4, $5)
       ON CONFLICT (user_id) 
       DO UPDATE SET shop_name = $2, currency_symbol = $3, invoice_counter = $4, low_stock_default = $5`,
      [currentUserId, settings.shopName, settings.currencySymbol, settings.invoiceCounter, settings.lowStockDefault]
    );
    return { success: true };
  } catch (err) {
    console.error("Save settings error:", err);
    throw new Error('Failed to save settings to database.');
  }
});

ipcMain.handle('db:send-invoice-email', async (event, invoice, user) => {
  const host = process.env.SMTP_HOST;
  const port = process.env.SMTP_PORT;
  const secure = process.env.SMTP_SECURE === 'true';
  const systemUser = process.env.SMTP_USER;
  const systemPass = process.env.SMTP_PASS;
  
  if (!host || !systemUser || !systemPass) {
    console.warn("SMTP credentials not configured in .env. Skipping customer invoice email dispatch.");
    return { success: false, error: 'SMTP credentials not configured.' };
  }

  try {
    const transporter = nodemailer.createTransport({
      host,
      port: parseInt(port) || 465,
      secure,
      auth: {
        user: systemUser,
        pass: systemPass
      }
    });

    const userEmail = user?.email || 'noreply@yourdomain.com';
    const userName = (user?.firstName && user?.lastName) ? `${user.firstName} ${user.lastName}` : 'Hardware Store';
    const orgName = user?.organizationName || 'Our Store';

    // From Display Name shows user's name and org name: "Alice Smith (Smith Tools) <systemMail>"
    const fromName = `${userName} (${orgName})`;
    const fromAddress = `"${fromName}" <${systemUser}>`;
    const replyToAddress = userEmail;

    // Build invoice items table rows
    let itemsHtml = '';
    const items = Array.isArray(invoice.items) ? invoice.items : [];
    for (const item of items) {
      const itemTotal = Number(item.quantity) * Number(item.price);
      itemsHtml += `
        <tr>
          <td style="padding: 10px; border-bottom: 1px solid #e0e0e0; font-family: sans-serif;">${item.name}</td>
          <td style="padding: 10px; border-bottom: 1px solid #e0e0e0; text-align: center; font-family: sans-serif;">${item.quantity} ${item.unit || 'pcs'}</td>
          <td style="padding: 10px; border-bottom: 1px solid #e0e0e0; text-align: right; font-family: sans-serif;">Rs ${Number(item.price).toFixed(2)}</td>
          <td style="padding: 10px; border-bottom: 1px solid #e0e0e0; text-align: right; font-weight: bold; font-family: sans-serif;">Rs ${itemTotal.toFixed(2)}</td>
        </tr>
      `;
    }

    const htmlContent = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 25px; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; color: #333333;">
        <div style="border-bottom: 2px solid #d9720b; padding-bottom: 15px; margin-bottom: 20px;">
          <h2 style="color: #d9720b; margin: 0; font-size: 22px; font-family: sans-serif;">INVOICE RECEIVED</h2>
          <p style="margin: 4px 0 0 0; font-size: 14px; color: #666666; font-family: sans-serif;">Thank you for your business with <strong>${orgName}</strong></p>
        </div>

        <table style="width: 100%; margin-bottom: 20px; font-size: 13px; color: #555555; border-collapse: collapse;">
          <tr>
            <td style="font-family: sans-serif; padding: 2px 0;"><strong>Invoice Number:</strong> ${invoice.invoiceNumber}</td>
            <td style="text-align: right; font-family: sans-serif; padding: 2px 0;"><strong>Date:</strong> ${new Date(invoice.date).toLocaleDateString()}</td>
          </tr>
          <tr>
            <td style="font-family: sans-serif; padding: 2px 0;"><strong>Customer Name:</strong> ${invoice.customerName || 'Walk-in customer'}</td>
            <td style="text-align: right; font-family: sans-serif; padding: 2px 0;"><strong>Email:</strong> ${invoice.customerEmail}</td>
          </tr>
        </table>

        <table style="width: 100%; border-collapse: collapse; font-size: 14px; margin-bottom: 20px;">
          <thead>
            <tr style="background-color: #f7f5f0;">
              <th style="padding: 10px; text-align: left; border-bottom: 2px solid #d9720b; color: #746C5E; font-size: 11px; font-weight: 700; font-family: sans-serif; letter-spacing: 0.5px;">ITEM DESCRIPTION</th>
              <th style="padding: 10px; text-align: center; border-bottom: 2px solid #d9720b; color: #746C5E; font-size: 11px; font-weight: 700; font-family: sans-serif; letter-spacing: 0.5px;">QTY</th>
              <th style="padding: 10px; text-align: right; border-bottom: 2px solid #d9720b; color: #746C5E; font-size: 11px; font-weight: 700; font-family: sans-serif; letter-spacing: 0.5px;">UNIT PRICE</th>
              <th style="padding: 10px; text-align: right; border-bottom: 2px solid #d9720b; color: #746C5E; font-size: 11px; font-weight: 700; font-family: sans-serif; letter-spacing: 0.5px;">TOTAL</th>
            </tr>
          </thead>
          <tbody>
            ${itemsHtml}
          </tbody>
        </table>

        <div style="width: 60%; margin-left: auto; font-size: 14px; line-height: 1.6; margin-bottom: 20px;">
          <table style="width: 100%; border-collapse: collapse;">
            <tr>
              <td style="color: #666666; font-family: sans-serif; padding: 3px 0;">Subtotal:</td>
              <td style="font-weight: 600; text-align: right; font-family: sans-serif; padding: 3px 0;">Rs ${Number(invoice.subtotal).toFixed(2)}</td>
            </tr>
            ${invoice.discount > 0 ? `
            <tr>
              <td style="color: #a03333; font-family: sans-serif; padding: 3px 0;">Discount:</td>
              <td style="color: #a03333; font-weight: 600; text-align: right; font-family: sans-serif; padding: 3px 0;">- Rs ${Number(invoice.discount).toFixed(2)}</td>
            </tr>
            ` : ''}
            <tr style="border-top: 1px solid #d9720b;">
              <td style="font-weight: bold; color: #d9720b; font-family: sans-serif; padding: 8px 0; font-size: 16px;">Total Paid:</td>
              <td style="font-weight: bold; color: #d9720b; text-align: right; font-family: sans-serif; padding: 8px 0; font-size: 16px;">Rs ${Number(invoice.total).toFixed(2)}</td>
            </tr>
          </table>
        </div>

        <br/>
        <div style="background-color: #fcfbf7; border-left: 4px solid #d9720b; padding: 12px 15px; font-size: 13px; color: #666666; border-radius: 4px; line-height: 1.4;">
          <strong>Contact Information:</strong> If you have any questions regarding this invoice, please email <strong>${userName}</strong> directly at <a href="mailto:${userEmail}" style="color: #d9720b; text-decoration: none;">${userEmail}</a>.
        </div>

        <br/>
        <p style="font-size: 11px; color: #999999; text-align: center; border-top: 1px solid #eeeeee; padding-top: 15px; margin-bottom: 0; font-family: sans-serif;">
          Invoice processed and sent on behalf of ${orgName} via Hardware Inventory Portal.
        </p>
      </div>
    `;

    await transporter.sendMail({
      from: fromAddress,
      to: invoice.customerEmail,
      replyTo: replyToAddress,
      subject: `Invoice ${invoice.invoiceNumber} from ${orgName}`,
      html: htmlContent
    });

    console.log(`Invoice email sent successfully to ${invoice.customerEmail} on behalf of ${userEmail}`);
    return { success: true };
  } catch (err) {
    console.error("Failed to send invoice email:", err);
    return { success: false, error: err.message };
  }
});

// ==========================================
// WINDOW & APP LIFE CYCLE
// ==========================================

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
