# BCA Price Analyzer - Claude Code Guide

This guide provides Claude Code with context and rules to help maintain this car auction bidding tool.

## Key Information
- **Purpose**: Calculates optimal bid prices for cars from BCA auctions based on Coches.net market data with 30-40% profit margin.
- **Architecture**: Node.js/Express server (`server.js`) with an HTML/JS frontend in `public/`.
- **Primary Logic**: Scraping Coches.net using Direct SSR extraction with Apify fallback.

## Coding Standards
- **Language**: JavaScript (CommonJS).
- **Framework**: Express.js.
- **Scraping**: Always use the direct extraction from `window.__INITIAL_PROPS__` in `server.js` as the primary method.
- **Normalization**: Use the `normalizeStr` function to handle car make/model accents to ensure dictionary matches.
- **Parameters**: `FuelTypeIds%5B0%5D`: 1=Gasoline, 2=Diesel. `TransmissionTypeId%5B0%5D`: 1=Manual, 2=Auto.

## Useful Commands
- **Run Server**: `npm run dev` or `node server.js`
- **Port**: 3000
- **Excel Handling**: Uses `xlsx` library to parse BCA reports from `uploads/`.

## File Structure
- `server.js`: Core backend logic and scraping.
- `public/`: Web frontend components.
- `coches_makes_models.json`: ID dictionary for Coches.net.
- `uploads/`: Temporary storage for Excel files.
