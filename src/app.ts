// src/app.ts
import express from 'express';
import connectDB from './config/db';
import sheetRoutes from './api/v1/router/sheetRouter'

const app = express();

app.use(express.json());

connectDB();

app.use("/sheets", sheetRoutes);

export default app;
