const quickbooksData = {
  accounts: [
    { AccountID: "1000", Name: "Cash", Type: "Bank", Active: true },
    { AccountID: "4000", Name: "Sales Revenue", Type: "Income", Active: true },
    { AccountID: "5000", Name: "Cost of Goods Sold", Type: "COGS", Active: true },
    { AccountID: "6000", Name: "Salaries Expense", Type: "Expense", Active: true }
  ],
  vendors: [
    { VendorID: "V001", Name: "ABC Supplies", Email: "ap@abcsupplies.com", Balance: 0 },
    { VendorID: "V002", Name: "XYZ Consulting", Email: "billing@xyz.com", Balance: 0 }
  ],
  customers: [
    { CustomerID: "CU01", Name: "Globex Corp", Email: "ar@globex.com" },
    { CustomerID: "CU02", Name: "Initech Ltd", Email: "finance@initech.com" }
  ]
};

const xeroData = {
  accounts: [
    { AccountID: "X100", Name: "Xero Bank Account", Type: "Bank", Active: true },
    { AccountID: "X400", Name: "Xero Sales", Type: "Revenue", Active: true },
    { AccountID: "X500", Name: "Xero Expenses", Type: "Expense", Active: true }
  ],
  vendors: [
    { VendorID: "XV001", Name: "Xero Vendor One", Email: "vendor1@xero-demo.com", Balance: 0 },
    { VendorID: "XV002", Name: "Xero Vendor Two", Email: "vendor2@xero-demo.com", Balance: 0 }
  ],
  customers: [
    { CustomerID: "XCU01", Name: "Xero Customer One", Email: "customer1@xero-demo.com" },
    { CustomerID: "XCU02", Name: "Xero Customer Two", Email: "customer2@xero-demo.com" }
  ]
};

export default function handler(req, res) {
  const { platform = "qbo", type = "accounts" } = req.query;

  const source = platform === "xero" ? xeroData : quickbooksData;
  const data = source[type];

  if (!data) {
    return res.status(400).json({
      error: "Invalid master data type",
      allowedTypes: ["accounts", "vendors", "customers"]
    });
  }

  res.status(200).json({
    platform,
    type,
    count: data.length,
    data
  });
}