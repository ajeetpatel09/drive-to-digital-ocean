import { ResourceNotFoundError } from "../common/errors";
import { Request } from "express";
import * as XLSX from "xlsx";
import * as ExcelJS from "exceljs";
import axios from "axios";
import * as AWS from "aws-sdk";

interface RowData {
  [key: string]: any;
}
const DIGITAL_OCEAN_ENDPOINT = process.env.DIGITAL_OCEAN_ENDPOINT!;
const DIGITAL_OCEAN_REGION = process.env.DIGITAL_OCEAN_REGION!;
const DIGITAL_OCEAN_ACCESS_KEYID = process.env.DIGITAL_OCEAN_ACCESS_KEYID!;
const DIGITAL_OCEAN_SECRET_ACCESS_KEY =
  process.env.DIGITAL_OCEAN_SECRET_ACCESS_KEY!;
// Configure AWS SDK for DigitalOcean Spaces
const spacesEndpoint = new AWS.Endpoint(DIGITAL_OCEAN_ENDPOINT);
const s3 = new AWS.S3({
  endpoint: spacesEndpoint,
  region: DIGITAL_OCEAN_REGION,
  credentials: {
    accessKeyId: DIGITAL_OCEAN_ACCESS_KEYID,
    secretAccessKey: DIGITAL_OCEAN_SECRET_ACCESS_KEY,
  },
});

function extractDriveId(driveLink: string): string | null {
  const regex = /\/d\/([a-zA-Z0-9_-]+)/;
  const match = driveLink.match(regex);
  return match ? match[1] : null;
}

async function fetchImageFromGoogleDrive(fileId: string) {
  const url = `https://drive.google.com/uc?export=download&id=${fileId}`;

  const response = await axios.get(url, { responseType: "arraybuffer" });
  const imageBuffer = Buffer.from(response.data, "binary");

  return imageBuffer;
}

class ReadExcelService {
  async readExcel(req: Request) {
    if (!req.file) {
      throw new ResourceNotFoundError("File not found.");
    }
    console.log("File uploaded successfully");

    // Read the Excel file
    const workbook = XLSX.read(req.file.buffer, { type: "buffer" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet);
    console.log("Data extracted successfully", data);

    const updatedData: RowData[] = await Promise.all(
      data.map(async (row: RowData) => {
        let driveLink = row["drive_link"]; // Adjust column name as needed

        if (driveLink) {
          try {
            const fileId = extractDriveId(driveLink);
            if (!fileId) {
              return { ...row, RemoteURL: "" };
            }
            const buffer = await fetchImageFromGoogleDrive(fileId);
            console.log("Image downloaded successfully", fileId);

            // Upload the image to DigitalOcean Spaces
            const key = `images/${row.wine_id} ${Date.now()}.jpg`;
            const params: AWS.S3.PutObjectRequest = {
              Bucket: "winebro",
              Key: key,
              Body: buffer,
              ContentType: "image/jpeg",
              ACL: "public-read",
            };
            const uploadResult = await s3.upload(params).promise();
            const remoteUrl = uploadResult.Location;
            console.log(`Image uploaded successfully: ${remoteUrl}`);

            // Update row with remote URL
            return { ...row, RemoteURL: remoteUrl };
          } catch (error) {
            console.error(`Error processing ${driveLink}:`, error);
            return row;
          }
        } else {
          return row;
        }
      })
    );

    // Create a new Excel file with updated URLs
    const newWorkbook = new ExcelJS.Workbook();
    const newSheet = newWorkbook.addWorksheet("Updated Data");
    newSheet.columns = Object.keys(updatedData[0]).map((key) => ({
      header: key,
      key,
    }));

    // Write the updated data to the new sheet
    updatedData.forEach((row) => newSheet.addRow(row));
    const output_path = "./output.xlsx";
    await newWorkbook.xlsx.writeFile(output_path);
    console.log("Excel file created successfully");
    return updatedData;
  }
}

export default ReadExcelService;
