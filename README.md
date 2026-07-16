# overtime_bot

Node.js API project using Express + TypeScript.

Project structure follows route -> controller -> service -> model for easier scaling.

## Requirements

- Node.js 20+

## Install

```bash
npm install
```

## Environment

Create `.env` from `.env.example` and update values for your environment.

## Run in Development

```bash
npm run dev
```

## Build

```bash
npm run build
```

## Run Production Build

```bash
npm start
```

## Lint

```bash
npm run lint
```

## Format

```bash
npm run format
```

## API

- `GET /` returns health status and project info.
- `GET /api/hello` returns a greeting message.
- `GET /api/employees/:employeeId/overtimes` returns overtime entries for an employee.
- `GET /api/employees/:employeeId/overtimes/summary` returns aggregated overtime summary.

Example:

```bash
curl "http://localhost:3000/api/hello?name=Ice"
```

```bash
curl "http://localhost:3000/api/employees/EMP001/overtimes"
```

```bash
curl "http://localhost:3000/api/employees/EMP001/overtimes/summary"
```