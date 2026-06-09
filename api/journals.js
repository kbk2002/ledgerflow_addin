export default function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({
      error: "Method not allowed. Use POST."
    });
  }

  const { platform = "qbo", transactionType = "journal", entries = [] } = req.body || {};

  if (!Array.isArray(entries) || entries.length === 0) {
    return res.status(400).json({
      error: "No journal entries provided"
    });
  }

  res.status(200).json({
    success: true,
    platform,
    transactionType,
    postedCount: entries.length,
    message: `Demo ${transactionType} posted to ${platform}`
  });
}