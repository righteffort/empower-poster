import { defineConfig } from "wxt";

export default defineConfig({
  srcDir: "src",
  outDir: "dist",
  manifest: {
    key: "MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA3wvMuIeaTabnQ5yPtpTMJVAMh6BWFqUwJW/A8Y3cqud+OV5eZv8B8q3ZKD53Se81+m6KN1mGDjrV8dQQY+7I9UP4+3K3qFGizbivCVPTOSPI1cPF3zKB1NGS1Po5wuKGAaGdtTl5EiSYPoM+DKSbAf9VOI0x9BnK6Vrge8/KGMNJ87mRGaRKstgWpTqYiArK1vnrb9FHNMPyGgkteTyJaKZk6S8GzXk+9Z36NDJOFjJblHijoZgxnzz13LWf2c2UYcDFM8UqV1Bh0LQTKi7rGp/BRsl+Dg9BcwHc/BDu7A/rlDwAOzkm+1LcRN7H5sJ5ztNqVHmNvw+XMWO6ulUPRwIDAQAB",
    icons: {
      "16": "icon/16.png",
      "32": "icon/32.png",
      "48": "icon/48.png",
      "128": "icon/128.png",
    },
    permissions: ["storage", "webRequest"],
    host_permissions: ["https://*.empower-retirement.com/*"],
    optional_host_permissions: ["*://*/*"],
    action: {
      default_title: "Empower Poster",
    },
  },
});
