-- ==========================================
-- PostgreSQL Database Schema for Hardware Inventory App
-- Production-ready multi-tenant schema with Row Level Security (RLS)
-- ==========================================

-- Trigger function to automatically update updated_at columns
CREATE OR REPLACE FUNCTION public.update_updated_at_column()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = NOW();
    RETURN NEW;
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

-- 1. PRODUCTS TABLE
CREATE TABLE IF NOT EXISTS public.products (
    id TEXT NOT NULL,
    user_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    category TEXT NOT NULL DEFAULT 'General',
    unit TEXT NOT NULL DEFAULT 'piece',
    quantity NUMERIC NOT NULL DEFAULT 0,
    cost_price NUMERIC NOT NULL DEFAULT 0,
    markup NUMERIC NOT NULL DEFAULT 40,
    low_stock INTEGER NOT NULL DEFAULT 5,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    PRIMARY KEY (user_id, id)
);

-- Trigger for products updated_at
DROP TRIGGER IF EXISTS tr_products_updated_at ON public.products;
CREATE TRIGGER tr_products_updated_at
    BEFORE UPDATE ON public.products
    FOR EACH ROW
    EXECUTE FUNCTION public.update_updated_at_column();

-- 2. INVOICES TABLE
CREATE TABLE IF NOT EXISTS public.invoices (
    id TEXT NOT NULL,
    user_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
    invoice_number TEXT NOT NULL,
    customer_name TEXT DEFAULT '',
    date TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    subtotal NUMERIC NOT NULL DEFAULT 0,
    discount NUMERIC NOT NULL DEFAULT 0,
    total NUMERIC NOT NULL DEFAULT 0,
    total_cost NUMERIC NOT NULL DEFAULT 0,
    profit NUMERIC NOT NULL DEFAULT 0,
    items JSONB NOT NULL DEFAULT '[]'::jsonb,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    PRIMARY KEY (user_id, id)
);

-- 3. EXPENSES TABLE
CREATE TABLE IF NOT EXISTS public.expenses (
    id TEXT NOT NULL,
    user_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
    date DATE NOT NULL DEFAULT CURRENT_DATE,
    description TEXT NOT NULL,
    category TEXT NOT NULL DEFAULT 'General',
    amount NUMERIC NOT NULL DEFAULT 0,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    PRIMARY KEY (user_id, id)
);

-- 4. SETTINGS TABLE (One settings record per user)
CREATE TABLE IF NOT EXISTS public.settings (
    user_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
    shop_name TEXT NOT NULL DEFAULT 'My Hardware Shop',
    currency_symbol TEXT NOT NULL DEFAULT 'Rs ',
    invoice_counter INTEGER NOT NULL DEFAULT 1,
    low_stock_default INTEGER NOT NULL DEFAULT 5,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    PRIMARY KEY (user_id)
);

-- Trigger for settings updated_at
DROP TRIGGER IF EXISTS tr_settings_updated_at ON public.settings;
CREATE TRIGGER tr_settings_updated_at
    BEFORE UPDATE ON public.settings
    FOR EACH ROW
    EXECUTE FUNCTION public.update_updated_at_column();


-- ==========================================
-- ROW LEVEL SECURITY (RLS) POLICIES
-- ==========================================

-- Enable RLS on all tables
ALTER TABLE public.products ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.invoices ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.expenses ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.settings ENABLE ROW LEVEL SECURITY;

-- PRODUCTS POLICIES
DROP POLICY IF EXISTS "Enable all access for users to their own products" ON public.products;
CREATE POLICY "Enable all access for users to their own products" ON public.products
    FOR ALL
    TO authenticated
    USING (auth.uid() = user_id)
    WITH CHECK (auth.uid() = user_id);

-- INVOICES POLICIES
DROP POLICY IF EXISTS "Enable all access for users to their own invoices" ON public.invoices;
CREATE POLICY "Enable all access for users to their own invoices" ON public.invoices
    FOR ALL
    TO authenticated
    USING (auth.uid() = user_id)
    WITH CHECK (auth.uid() = user_id);

-- EXPENSES POLICIES
DROP POLICY IF EXISTS "Enable all access for users to their own expenses" ON public.expenses;
CREATE POLICY "Enable all access for users to their own expenses" ON public.expenses
    FOR ALL
    TO authenticated
    USING (auth.uid() = user_id)
    WITH CHECK (auth.uid() = user_id);

-- SETTINGS POLICIES
DROP POLICY IF EXISTS "Enable all access for users to their own settings" ON public.settings;
CREATE POLICY "Enable all access for users to their own settings" ON public.settings
    FOR ALL
    TO authenticated
    USING (auth.uid() = user_id)
    WITH CHECK (auth.uid() = user_id);


-- ==========================================
-- PERFORMANCE INDEXES
-- ==========================================
CREATE INDEX IF NOT EXISTS idx_products_user_id ON public.products(user_id);
CREATE INDEX IF NOT EXISTS idx_invoices_user_id ON public.invoices(user_id);
CREATE INDEX IF NOT EXISTS idx_expenses_user_id ON public.expenses(user_id);
