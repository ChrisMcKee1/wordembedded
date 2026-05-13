import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  images: {
    remotePatterns: [
      {
        protocol: "https",
        hostname: "**.sharepoint.com",
      },
      {
        protocol: "https",
        hostname: "**.microsoft.com",
      },
    ],
  },
  transpilePackages: ["@fluentui/react-components"],
};

export default nextConfig;
