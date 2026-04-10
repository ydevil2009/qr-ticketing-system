# Free setup: Cloudinary images + Supabase Excel/database

## 1) Supabase
- Create a free Supabase project.
- Open SQL Editor and run `supabase-schema.sql`.
- In Storage, create a bucket named `exports`.
- Make the `exports` bucket public.
- In Project Settings -> API, copy:
  - Project URL
  - service_role key

## 2) Cloudinary
- Create a free Cloudinary account.
- From Dashboard copy:
  - Cloud name
  - API key
  - API secret

## 3) Brevo
- Create/keep a free Brevo account.
- Verify your sender email or domain.
- Copy the API key.

## 4) Local env
- Copy `.env.example` to `.env`
- Fill all values.

## 5) Install and run locally
- `npm install`
- `npm start`
- Open `http://localhost:5500`

## 6) Deploy on Render free web service
- Push code to GitHub.
- Create a new Web Service on Render.
- Build command: `npm install`
- Start command: `npm start`
- Add all variables from `.env.example` into Render environment variables.
- Set `BASE_URL` to your Render URL, for example: `https://your-app.onrender.com`

## Notes
- Payment screenshots go to Cloudinary.
- Ticket records go to Supabase Postgres.
- The latest Excel file goes to Supabase Storage bucket `exports` and is also downloadable from `/download-excel`.
- Never expose `SUPABASE_SERVICE_ROLE_KEY` on frontend/client code.
