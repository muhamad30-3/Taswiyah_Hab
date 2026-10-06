CREATE INDEX IF NOT EXISTS app_users_created_at_idx
  ON public.app_users (created_at DESC);

CREATE INDEX IF NOT EXISTS user_files_owner_uploaded_at_idx
  ON public.user_files (owner, uploaded_at DESC);
