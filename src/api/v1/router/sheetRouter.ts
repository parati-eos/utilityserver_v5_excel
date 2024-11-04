// src/routes/sheetsRouter.ts
import express from "express";
import { appendToSheet } from "../controller/sheetController";

const router = express.Router();

router.post("/append", appendToSheet);

export default router;
