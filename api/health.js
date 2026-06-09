export default function handler(req, res) {
  res.status(200).json({
    status: "ok",
    app: "LedgerFlow API",
    message: "Backend is running on Vercel"
  });
}