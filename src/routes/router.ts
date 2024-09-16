import { Router } from "express";
import ReadExcelController from "../controllers/readExcel.controller";
import * as multer from "multer";

const storage = multer.memoryStorage();
const upload = multer({ storage });

const router = Router({ mergeParams: true });
const readExcelController = new ReadExcelController();

router.post("/readExcel", upload.single("file"), readExcelController.readExcel);

export default router;
