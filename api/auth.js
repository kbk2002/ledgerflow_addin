export default function handler(req, res) {
  const { platform = "qbo" } = req.query;

  if (!["qbo", "xero"].includes(platform)) {
    return res.status(400).json({
      error: "Invalid platform. Use qbo or xero."
    });
  }

  res.status(200).json({
    status: "demo_connected",
    platform,
    companyName: platform === "qbo" ? "Acme Corp (QBO)" : "Acme Corp (Xero)",
    accessToken: "DEMO_TOKEN"
  });
}