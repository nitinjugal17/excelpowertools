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
      // Use an array of glob patterns to explicitly ignore directories.
      // This prevents Next.js from watching these folders for changes.
      config.watchOptions.ignored = [
        '**/.git/**',
        '**/node_modules/**',
        '**/.next/**',
        '**/uploads/**',
        '**/src/data/**', // Ignore the data directory to prevent restarts on settings save
      ];
    }

    return config;
  },
};

module.exports = nextConfig;
