/** @type {import('next').NextConfig} */
const nextConfig = {
  /* config options here */
  typescript: {
    ignoreBuildErrors: true,
  },
  eslint: {
    ignoreDuringBuilds: true,
  },
  images: {
    remotePatterns: [
      {
        protocol: 'https',
        hostname: 'placehold.co',
        port: '',
        pathname: '/**',
      },
    ],
  },
  webpack: (config, { isServer, dev }) => {
    if (isServer) {
      // These packages are not compatible with the server build, so we mark them as external.
      config.externals.push('xlsx-js-style', '@opentelemetry/instrumentation', 'jspdf', 'jspdf-autotable', 'handlebars', '@opentelemetry/sdk-node');
    }

    // In development, ignore specific directories to prevent server restarts.
    if (dev) {
      // Use an array to explicitly ignore dotfiles, node_modules, and our uploads directory.
      // This prevents Next.js from reloading when files change in these folders.
      config.watchOptions.ignored = [
        /(^|[\/\\])\../, // Default Next.js rule for dotfiles
        /node_modules/, // Default Next.js rule for node_modules
        /uploads/, // Custom rule for our uploads directory
      ];
    }

    return config;
  },
};

module.exports = nextConfig;
