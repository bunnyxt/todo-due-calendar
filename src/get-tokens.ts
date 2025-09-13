import "dotenv/config";
import fetch from "node-fetch";

const CLIENT_ID = process.env.CLIENT_ID;

if (!CLIENT_ID) {
  console.error("❌ Please set CLIENT_ID in your .env file");
  console.log("\nTo get CLIENT_ID:");
  console.log("1. Go to https://portal.azure.com");
  console.log("2. Navigate to Azure Active Directory > App registrations");
  console.log("3. Create a new registration or use an existing one");
  console.log("4. Copy the Application (client) ID");
  process.exit(1);
}

async function getTokensWithDeviceFlow() {
  console.log("🔑 Starting device code flow...");

  // Step 1: Get device code
  const deviceUrl =
    "https://login.microsoftonline.com/common/oauth2/v2.0/devicecode";
  const deviceData = new URLSearchParams({
    client_id: CLIENT_ID!,
    scope: "https://graph.microsoft.com/Tasks.Read offline_access",
  });

  const deviceResponse = await fetch(deviceUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: deviceData,
  });

  if (!deviceResponse.ok) {
    const errorText = await deviceResponse.text();
    throw new Error(
      `Device code request failed: ${deviceResponse.status} ${errorText}`
    );
  }

  const deviceInfo = (await deviceResponse.json()) as {
    device_code: string;
    user_code: string;
    verification_uri: string;
    verification_uri_complete: string;
    expires_in: number;
    interval: number;
  };

  console.log("\n📱 Please complete the following steps:");
  console.log(`1. Open this URL: ${deviceInfo.verification_uri}`);
  console.log(`2. Enter this code: ${deviceInfo.user_code}`);
  console.log(
    `3. Or click this direct link: ${deviceInfo.verification_uri_complete}`
  );
  console.log("\n⏳ Waiting for you to complete the authentication...");

  // Step 2: Poll for tokens
  const tokenUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
  const tokenData = new URLSearchParams({
    grant_type: "urn:ietf:params:oauth:grant-type:device_code",
    client_id: CLIENT_ID!,
    device_code: deviceInfo.device_code,
  });

  const startTime = Date.now();
  const expiryTime = startTime + deviceInfo.expires_in * 1000;

  while (Date.now() < expiryTime) {
    await new Promise((resolve) =>
      setTimeout(resolve, deviceInfo.interval * 1000)
    );

    const tokenResponse = await fetch(tokenUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: tokenData,
    });

    const tokenResult = (await tokenResponse.json()) as {
      access_token?: string;
      refresh_token?: string;
      expires_in?: number;
      scope?: string;
      error?: string;
      error_description?: string;
    };

    if (tokenResponse.ok) {
      const tokens = tokenResult as {
        access_token: string;
        refresh_token: string;
        expires_in: number;
        scope: string;
      };

      console.log("\n✅ Tokens retrieved successfully!");
      console.log("📝 Add these to your .env file:");
      console.log(`ACCESS_TOKEN=${tokens.access_token}`);
      console.log(`REFRESH_TOKEN=${tokens.refresh_token}`);
      console.log(`CLIENT_ID=${CLIENT_ID}`);

      return;
    } else if (tokenResult.error === "authorization_pending") {
      process.stdout.write(".");
      continue;
    } else if (tokenResult.error === "authorization_declined") {
      throw new Error("Authorization was declined by the user");
    } else if (tokenResult.error === "expired_token") {
      throw new Error("Device code expired. Please try again.");
    } else {
      throw new Error(
        `Token request failed: ${tokenResult.error} - ${tokenResult.error_description}`
      );
    }
  }

  throw new Error("Device code expired. Please try again.");
}

getTokensWithDeviceFlow().catch((error) => {
  console.error("❌ Error:", error.message);
  process.exit(1);
});
