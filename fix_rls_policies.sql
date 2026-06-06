-- ============================================================
-- UniTrack v3.0 — RLS Fix for Anonymous Access
-- ============================================================
-- PROBLEM: The app connects as 'anon' (no login), but RLS
-- policies only allow 'authenticated'. This silently blocks
-- all inserts/updates/deletes.
--
-- HOW TO RUN: Paste this into Supabase Dashboard → SQL Editor → Run
-- ============================================================

-- 1. Add anonymous SELECT policy for employees
CREATE POLICY "Allow anon select on employees"
ON public.employees FOR SELECT TO anon USING (true);

-- 2. Add anonymous INSERT policy for employees
CREATE POLICY "Allow anon insert on employees"
ON public.employees FOR INSERT TO anon WITH CHECK (true);

-- 3. Add anonymous UPDATE policy for employees
CREATE POLICY "Allow anon update on employees"
ON public.employees FOR UPDATE TO anon USING (true) WITH CHECK (true);

-- 4. Add anonymous DELETE policy for employees
CREATE POLICY "Allow anon delete on employees"
ON public.employees FOR DELETE TO anon USING (true);

-- 5. Add anonymous SELECT policy for transactions
CREATE POLICY "Allow anon select on transactions"
ON public.transactions FOR SELECT TO anon USING (true);

-- 6. Add anonymous INSERT policy for transactions
CREATE POLICY "Allow anon insert on transactions"
ON public.transactions FOR INSERT TO anon WITH CHECK (true);

-- 7. Add anonymous UPDATE policy for transactions
CREATE POLICY "Allow anon update on transactions"
ON public.transactions FOR UPDATE TO anon USING (true) WITH CHECK (true);

-- 8. Add anonymous DELETE policy for transactions
CREATE POLICY "Allow anon delete on transactions"
ON public.transactions FOR DELETE TO anon USING (true);

-- 9. Add anonymous SELECT policy for config
CREATE POLICY "Allow anon select on config"
ON public.config FOR SELECT TO anon USING (true);

-- 10. Add anonymous INSERT policy for config
CREATE POLICY "Allow anon insert on config"
ON public.config FOR INSERT TO anon WITH CHECK (true);

-- 11. Add anonymous UPDATE policy for config
CREATE POLICY "Allow anon update on config"
ON public.config FOR UPDATE TO anon USING (true) WITH CHECK (true);

-- 12. Add anonymous DELETE policy for config
CREATE POLICY "Allow anon delete on config"
ON public.config FOR DELETE TO anon USING (true);

-- ============================================================
-- DONE! The app should now be able to read/write all tables.
-- ============================================================
