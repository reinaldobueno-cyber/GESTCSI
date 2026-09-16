CREATE TABLE IF NOT EXISTS sessions (
  session_hash TEXT PRIMARY KEY,
  token_cipher TEXT NOT NULL,
  user_json TEXT NOT NULL,
  csrf_hash TEXT NOT NULL,
  created_at INTEGER NOT NULL,
  expires_at INTEGER NOT NULL
);

CREATE INDEX IF NOT EXISTS sessions_expires_at ON sessions (expires_at);

CREATE TABLE IF NOT EXISTS login_attempts (
  attempt_key TEXT PRIMARY KEY,
  window_start INTEGER NOT NULL,
  attempts INTEGER NOT NULL
);

CREATE INDEX IF NOT EXISTS login_attempts_window_start ON login_attempts (window_start);
