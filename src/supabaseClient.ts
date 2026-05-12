import { createClient } from '@supabase/supabase-js'

// Credenciais do seu projeto Supabase
// Nota: Para produção, recomenda-se usar variáveis de ambiente (.env)
const supabaseUrl = 'https://ggzthmcznjlyvxchpfzh.supabase.co'
const supabaseAnonKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImdnenRobWN6bmpseXZ4Y2hwZnpoIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg1ODI1NjYsImV4cCI6MjA5NDE1ODU2Nn0.0D1r07TJFvQmySYob-tfwJeJLQmRVqWzpKXOFRJnlXE'

export const supabase = createClient(supabaseUrl, supabaseAnonKey)
