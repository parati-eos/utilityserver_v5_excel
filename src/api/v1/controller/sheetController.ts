import { Request, Response } from "express";
import { writeToSheets } from "../../../services/dataflowservice";

export const appendToSheet = async (req: Request, res: Response): Promise<void> => {
    const jsonData = req.body;
    const spreadsheetId = process.env.SHEET_ID || "1QKrcz69mF2xsuoYlsWCov9IQBwSmYThv8QxmD4oAdcA";

    if (!jsonData || typeof jsonData !== "object" || Array.isArray(jsonData)) {
        res.status(400).json({ error: "Invalid data format: jsonData should be a non-empty MongoDB JSON object" });
        return;
    }

    try {
        // Convert MongoDB JSON into an array format compatible with Google Sheets
        // const sheetData = Object.keys(jsonData).map((key) => [key, jsonData[key]]);

        const result = await writeToSheets(jsonData);
        res.status(200).json({ message: "Data appended successfully", result });
    } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "An unexpected error occurred";
        res.status(500).json({ error: errorMessage });
    }
};
