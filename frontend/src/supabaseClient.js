import { createClient } from '@supabase/supabase-js'

const supabaseUrl = 'https://pckwvybcvabjsvlygqtj.supabase.co'
const supabaseAnonKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBja3d2eWJjdmFianN2bHlncXRqIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODA0NzI1MDksImV4cCI6MjA5NjA0ODUwOX0.6pOmSJ_divtragSkwvotDem23lrIF8rfvJILoyIGrr8'

export const supabase = createClient(supabaseUrl, supabaseAnonKey)
