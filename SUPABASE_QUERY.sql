-- ====================================================================
-- MCO MAKER - SUPABASE QUERIES
-- Multi-Tenant Schema & Reporting for Supabase SQL Editor
-- ====================================================================

-- ====================================================================
-- TAB 1: INITIAL SETUP SCRIPT (Run Once in Supabase SQL Editor)
-- ====================================================================

-- 1. Ensure required extensions are available
CREATE EXTENSION IF NOT EXISTS "pgcrypto";

-- 2. Table: mco_licenses (Serial keys & subscriptions)
CREATE TABLE IF NOT EXISTS public.mco_licenses (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    serial_key TEXT UNIQUE NOT NULL,
    plan_type TEXT NOT NULL CHECK (plan_type IN ('6_months', '1_year')),
    duration_days INTEGER NOT NULL CHECK (duration_days IN (180, 365)),
    max_devices INTEGER NOT NULL CHECK (max_devices IN (1, 2, 3)),
    activated_at TIMESTAMPTZ,
    expires_at TIMESTAMPTZ,
    is_revoked BOOLEAN NOT NULL DEFAULT FALSE,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 3. Table: mco_license_devices (Registered user computers)
CREATE TABLE IF NOT EXISTS public.mco_license_devices (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    license_id UUID NOT NULL REFERENCES public.mco_licenses(id) ON DELETE CASCADE,
    device_id TEXT NOT NULL,
    os_info TEXT,
    first_seen TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    last_active TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    CONSTRAINT uq_mco_license_device UNIQUE (license_id, device_id)
);

-- 4. Table: mco_app_updates (Auto-updater release notices)
CREATE TABLE IF NOT EXISTS public.mco_app_updates (
    version VARCHAR(32) PRIMARY KEY,
    download_url TEXT NOT NULL,
    file_name VARCHAR(128) DEFAULT 'MCOMaker-Setup.exe',
    file_size BIGINT DEFAULT 0,
    release_notes TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- 5. Performance Indexes
CREATE INDEX IF NOT EXISTS idx_mco_licenses_serial_key ON public.mco_licenses (serial_key);
CREATE INDEX IF NOT EXISTS idx_mco_licenses_expires_at ON public.mco_licenses (expires_at);
CREATE INDEX IF NOT EXISTS idx_mco_license_devices_license_id ON public.mco_license_devices (license_id);
CREATE INDEX IF NOT EXISTS idx_mco_license_devices_device_id ON public.mco_license_devices (device_id);

-- 6. Row Level Security (RLS) Configuration
ALTER TABLE public.mco_licenses ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.mco_license_devices ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.mco_app_updates ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "Deny anon access on mco_licenses" ON public.mco_licenses;
CREATE POLICY "Deny anon access on mco_licenses" ON public.mco_licenses
    FOR ALL TO anon USING (false);

DROP POLICY IF EXISTS "Deny anon access on mco_license_devices" ON public.mco_license_devices;
CREATE POLICY "Deny anon access on mco_license_devices" ON public.mco_license_devices
    FOR ALL TO anon USING (false);

DROP POLICY IF EXISTS "Allow anon read on mco_app_updates" ON public.mco_app_updates;
CREATE POLICY "Allow anon read on mco_app_updates" ON public.mco_app_updates
    FOR SELECT TO anon USING (true);

DROP POLICY IF EXISTS "Deny anon modify on mco_app_updates" ON public.mco_app_updates;
CREATE POLICY "Deny anon modify on mco_app_updates" ON public.mco_app_updates
    FOR INSERT TO anon WITH CHECK (false);

COMMENT ON TABLE public.mco_licenses IS 'Serial keys, plans, and activation lifecycles for MCO Maker.';
COMMENT ON TABLE public.mco_license_devices IS 'Hardware devices bound to active licenses with device limits.';
COMMENT ON TABLE public.mco_app_updates IS 'Release versions and download pointers for MCO Maker.';


-- ====================================================================
-- TAB 2: LIVE USERS MONITORING QUERY (Save in separate tab 'Users MCOMaker')
-- ====================================================================
/*
SELECT 
    'MCO Maker' AS app_name,
    l.serial_key,
    l.plan_type,
    d.device_id AS computer_hwid,
    d.os_info AS windows_version,
    d.first_seen AS activated_on,
    d.last_active AS last_used_on,
    l.expires_at,
    l.max_devices,
    l.is_revoked
FROM public.mco_license_devices d
JOIN public.mco_licenses l ON d.license_id = l.id
ORDER BY d.last_active DESC;
*/
