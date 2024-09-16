import { Request, Response } from "express";
import ReadExcelService from "../services/readExcel.service";
import { buildResponse } from "../common/utils";

const readExcelService = new ReadExcelService();
class ReadExcelController {
  async readExcel(req: Request, res: Response) {
    const result = await readExcelService.readExcel(req);

    return res.status(200).send(buildResponse(result, "success", null));
  }
}

export default ReadExcelController;
