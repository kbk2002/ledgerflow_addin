import { createClient } from "@supabase/supabase-js";

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

export default async function handler(req, res) {

  if (req.method !== "POST") {
    return res.status(405).json({
      error: "Method not allowed"
    });
  }

  try {

    const { mappings } = req.body;

    const { error } = await supabase
      .from("account_mappings")
      .insert(mappings);

    if (error) throw error;

    return res.status(200).json({
      success: true,
      count: mappings.length
    });

  } catch (err) {

    return res.status(500).json({
      error: err.message
    });
  }
}
