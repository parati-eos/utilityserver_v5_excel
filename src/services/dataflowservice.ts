import { google } from "googleapis";
import * as fs from "fs";
import { sheets } from "googleapis/build/src/apis/sheets";

// Load Google service account credentials
// const credentials = JSON.parse(fs.readFileSync('config/credentials.json', 'utf-8'));

const credentials = {
  type: "service_account",
  project_id: "presentation1-408602",
  private_key_id: "9598c7d8770f74d25818896fca2d1b0e34c54d9f",
  private_key:
    "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQC7LKKD6Rm5oPp+\nhB7jxGjr2J730LozdFv9/75PYNHXKW57VJPKkMP7Lx7PJWXBTrmVcn29fLKHuPQB\nSpI4aB7awl2s5a2HudFgUlitj/73QVm4E8PeVCu00f+fYDqFfcyx9TLBGpLrJKeg\nxjz8wgAtpH/SjVLa3pwlBpSQjMtDK1uLd7uoEJ+fgMZGaYXAXEz5KL+ZbeVsXGKM\nLCndcCGVuDW1s3ilqmRlKGsWs7WBUoZPvfMi9Xmi/dOU/hP1TQHUZEFdBog+WcIt\n7swxuoFKnoFIl77oWG7C578h9z3UZQdV2xUcR61cQ/UxSB/IBsZR6TO9wv8whDuX\nJqcTX493AgMBAAECggEAAXzmUCV71UE3sjkfa6yb1QnMoXKFD6F0+HWqwpZjzGJB\niyiIw/enqskc6YUnCVdyahe3GAp0G/UfVLm/nBKI2SiXagpvz+2eqxzEuRz/eKOz\nYgqjmbN9lr5SSY49sRwpJ26sQ2NKHDXVKFPk5kgNCC7R3lXFx2ugkLZlY2vfOkoH\n6R2xtsDs2dBTdQoK3MJ4ii1Lr/hbFETv1EFHhn0avDkQnGZs361ww4n1IcmpeDAT\ngDPj8Cxi1j0xiUED/dZTUhaZbmmEUj/TssPjhQaNnbsSkI2/zk6iQhSjfR8eA/Yi\ntufYcNuj5uksC0EsMnEPcv/KgsW73gcpfPQcPJGcYQKBgQD9rRAol1Ni/XFWm7gS\nuUyO/FZMopPb52ARmkDamxKfCtI3kEgL5McsPiQUWc4ZeELJicSFwGXe0xVUuPxy\nSJb6westwoxsBVhpclwxGntc8GfU4uRxlcuzMGGCxUnum1s+5ukEtAUqvOl9g5gJ\n/QrkI+I2efSpzDYBQMdyyKemSQKBgQC845uaY/5j5f3h4wVyWR5yoPcBctoPfU5A\nUVa7Qf4Nw1UdCLHogJkyyfapiz2Z27pTGQSMlLbEu6q44KYw5qKAGa1WQLxruKm2\n0ihjak3KVeFwH6xXR/eD5mGj9S/HQihkiJJTZfgUP3aweG6CBom2Rh1Ddpm1qdTK\nVombwMqHvwKBgBUuPgsll3DMeIoitlvZ3OqTZyE+8dmKmBrgJkoaaJOe865v/ZQA\npiCrj5ejZ/H4eJsbRa1lQxw3w7AvQeTI6tJFHr3TYKYkTB2BzvDKpUI9UG4WA7z4\nJOnxQDMLBgFGN3gpD4u0/Dl1TImOU0OCPUaPOHQT+rmys0+neP+8gUMBAoGARBUu\nEtoT6WIOvoqbffnNVbfbEDSbkJWzzM8Emf5RWhib5xkpNwqTLZFKTRYZIAnpAOa1\nkw5PSl3yTSz7+ghHbjDTH5G52IH4+iKJ2DuKynFmDon8DoGsH2i8rOJFVGbuND5d\nr53Da1jsqPLfshI1NPPUvGpVQPtz7XJ/qxo0ZfUCgYEA3Y2u+EikXGYbZG4YUKDd\nhZn9wa+mjEknqCVMe4a6DmH4NBJABAJhBp3FXtV9XJ/O1AAY3cES6kq16Ny0UXmh\nSn+m1eikJGlRju5mZpADGEeu396tCVIJFzGJoXr+vjEZ0DwdnL6rTrHBOHFY4Q+e\njLQRvgFA9yhXaFvdeImXThk=\n-----END PRIVATE KEY-----\n",
  client_email: "form-pitch@presentation1-408602.iam.gserviceaccount.com",
  client_id: "116193399203206103850",
  auth_uri: "https://accounts.google.com/o/oauth2/auth",
  token_uri: "https://oauth2.googleapis.com/token",
  auth_provider_x509_cert_url: "https://www.googleapis.com/oauth2/v1/certs",
  client_x509_cert_url:
    "https://www.googleapis.com/robot/v1/metadata/x509/form-pitch%40presentation1-408602.iam.gserviceaccount.com",
  universe_domain: "googleapis.com",
};

// Google Spreadsheet ID and sheet names
const SPREADSHEET_ID = "1QKrcz69mF2xsuoYlsWCov9IQBwSmYThv8QxmD4oAdcA";
const SHEETS = ["Cover", "Points", "Phases", "People"] as const;

// Authorize and get Sheets client
async function authorize() {
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  return google.sheets({ version: "v4", auth });
}

// Flatten nested objects for simpler data rows
function flattenObject(ob: any): any {
  const toReturn: Record<string, any> = {};
  for (const i in ob) {
    if (!ob.hasOwnProperty(i)) continue;
    if (typeof ob[i] === "object" && ob[i] !== null && !Array.isArray(ob[i])) {
      const flatObject = flattenObject(ob[i]);
      for (const x in flatObject) {
        if (!flatObject.hasOwnProperty(x)) continue;
        toReturn[i + "." + x] = flatObject[x];
      }
    } else {
      toReturn[i] = ob[i];
    }
  }
  return toReturn;
}

// Prepare data for each sheet
function prepareSheetData(data: any) {
  return data.map((item: any) => Object.values(flattenObject(item)));
}

// Main function to write data to Google Sheets
export async function writeToSheets(jsonData: any) {
  const sheetsClient = await authorize();

  ////////////////////////////////////////////////////////////////////////////////////////////////////////////////

  let linearcover: any = [];
  linearcover = [
    "jsonData.userID",
    jsonData.documentID,
    jsonData.cover.companyName,
    "jsonData.cover.slideName",
    "jsonData.cover.designType",
    jsonData.cover.logo,
    jsonData.cover.tagline,
    jsonData.cover.image[0],
    "json.color.P100",
    "json.color.P75 - S25",
    "json.color.P50 - S50",
    "json.color.P25 - S75",
    "json.color.S100",
    "json.color.F - P100",
    "json.color.F - P75 - S25",
    "json.color.F - P50 - S50",
    "json.color.F - P25 - S75",
    "json.color.F - S100",
    "json.color.SCL",
    "json.color.SCD",
  ];

  try {
    await sheetsClient.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "Cover",
      valueInputOption: "RAW",
      requestBody: {
        values: [linearcover],
      },
    });

    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    const points = jsonData["points"];
    let linearpoints: any[] = [];

    for (let point of points) {
      linearpoints.push([
        "jsonData.userID",
        jsonData.documentID,
        "jsonData.companyName",
        point.slideName,
        "point.designType",
        "point.title",
        point.overview,
        point.header1,
        point.header2,
        point.header3,
        point.header4,
        point.header5,
        point.header6,
        point.description1,
        point.description2,
        point.description3,
        point.description4,
        point.description5,
        point.description6,
        point.icon1,
        point.icon2,
        point.icon3,
        point.icon4,
        point.icon5,
        point.icon6,
        point.image,
      ]);
    }

    await sheetsClient.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "Points",
      valueInputOption: "RAW",
      requestBody: { values: linearpoints },
    });

    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    const phases = jsonData["phases"];
    let linearphases: any[] = [];

    for (let phase of phases) {
      linearphases.push([
        "json.userID",
        jsonData.documentID,
        "json.companyName",
        phase.slideName,
        "phase.designType",
        "phase.title",
        phase.overview,
        phase.timeline1,
        phase.timeline2,
        phase.timeline3,
        phase.timeline4,
        phase.timeline5,
        phase.timeline6,
        phase.header1,
        phase.header2,
        phase.header3,
        phase.header4,
        phase.header5,
        phase.header6,
        phase.description1,
        phase.description2,
        phase.description3,
        phase.description4,
        phase.description5,
        phase.description6,
        phase.icon1,
        phase.icon2,
        phase.icon3,
        phase.icon4,
        phase.icon5,
        phase.icon6,
        phase.image,
      ]);
    }
    await sheetsClient.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "Phases",
      valueInputOption: "RAW",
      requestBody: { values: linearphases },
    });

    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    const images = jsonData["images"];
    let linearimages: any[] = [];

    for (let image of images) {
      linearimages.push([
        "userID",
        jsonData.documentID,
        jsonData.cover.companyName,
        image.slideName,
        image.designType,
        image.title,
        image.imageUrl1,
        image.imageUrl2,
        image.imageUrl3,
        image.imageUrl4,
        image.imageUrl5,
        image.imageUrl6,
        image.description1,
        image.description2,
        image.description3,
        image.description4,
        image.description5,
        image.description6,
      ]);
    }
    await sheetsClient.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "Images",
      valueInputOption: "RAW",
      requestBody: { values: linearimages },
    });

    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    //     const statisticss = jsonData["statisticss"];
    //     let linearstatisticss: any[] = [];

    //     for (let stat of statisticss) {
    //       linearstatisticss.push([
    //         "userID",
    //         jsonData.documentID,
    //         jsonData.cover.companyName,
    //         stat.slideName,
    //         stat.designType,
    //         stat.title,
    //         stat.overview,
    //         stat.stat1,
    //         stat.stat2,
    //         stat.stat3,
    //         stat.stat4,
    //         stat.stat5,
    //         stat.stat6,
    //         stat.header1,
    //         stat.header2,
    //         stat.header3,
    //         stat.header4,
    //         stat.header5,
    //         stat.header6,
    //         stat.description1,
    //         stat.description2,
    //         stat.description3,
    //         stat.description4,
    //         stat.description5,
    //         stat.description6,
    //         stat.icon1,
    //         stat.icon2,
    //         stat.icon3,
    //         stat.icon4,
    //         stat.icon5,
    //         stat.icon6,
    //         stat.image,
    //       ]);
    //     }
    //     await sheetsClient.spreadsheets.values.append({
    //       spreadsheetId: SPREADSHEET_ID,
    //       range: "Statistics",
    //       valueInputOption: "RAW",
    //       requestBody: { values: linearstatisticss },
    // });

    ////////////////////////////////////////////////////////////////////////////////////
    const people = jsonData["people"];
    let linearpeople: any[] = [];

    for (let person of people) {
      linearpeople.push([
        "userID",
        jsonData.documentID,
        jsonData.cover.companyName,
        person.slideName,
        person.designType,
        person.title,
        person.overview,
        person.name1,
        person.name2,
        person.name3,
        person.name4,
        person.name5,
        person.name6,
        person.designation1,
        person.designation2,
        person.designation3,
        person.designation4,
        person.designation5,
        person.designation6,
        person.description1,
        person.description2,
        person.description3,
        person.description4,
        person.description5,
        person.description6,
        person.personUrl1,
        person.personUrl2,
        person.personUrl3,
        person.personUrl4,
        person.personUrl5,
        person.personUrl6,
        person.image,
      ]);
    }
    await sheetsClient.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "People",
      valueInputOption: "RAW",
      requestBody: { values: linearpeople },
    });

    ////////////////////////////////////////////////////////////////////////////////////
    const tables = jsonData["tables"];
    let lineartables: any[] = [];

    for (let table of tables) {
      console.log(table)
      lineartables.push([
        "userID",
        jsonData.documentID,
        jsonData.cover.companyName,
        table.slideName,
        table.designType,
        table.title,
        table.overview,
        table.rowHeader1,
        table.rowHeader2,
        table.rowHeader3,
        table.rowHeader4,
        table.rowHeader5,
        table.rowHeader6,
        table.rowHeader7,
        table.rowHeader8,
        table.columnHeader1,
        table.columnHeader2,
        table.columnHeader3,
        table.columnHeader4,
        table.columnHeader5,
        table.row1.attribute1,
        table.row1.attribute2,
        table.row1.attribute3,
        table.row1.attribute4,
        table.row1.attribute5,
        table.row2.attribute1,
        table.row2.attribute2,
        table.row2.attribute3,
        table.row2.attribute4,
        table.row2.attribute5,
        table.row3.attribute1,
        table.row3.attribute2,
        table.row3.attribute3,
        table.row3.attribute4,
        table.row3.attribute5,
        table.row4.attribute1,
        table.row4.attribute2,
        table.row4.attribute3,
        table.row4.attribute4,
        table.row4.attribute5,
        table.row5.attribute1,
        table.row5.attribute2,
        table.row5.attribute3,
        table.row5.attribute4,
        table.row5.attribute5,
        table.row6.attribute1,
        table.row6.attribute2,
        table.row6.attribute3,
        table.row6.attribute4,
        table.row6.attribute5,
        table.row7.attribute1,
        table.row7.attribute2,
        table.row7.attribute3,
        table.row7.attribute4,
        table.row7.attribute5,
        table.row8.attribute1,
        table.row8.attribute2,
        table.row8.attribute3,
        table.row8.attribute4,
        table.row8.attribute5,
      ]);
    }
    await sheetsClient.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "Tables",
      valueInputOption: "RAW",
      requestBody: { values: lineartables },
    });

    console.log("Sheets updated successfully.");
  } catch (error) {
    console.error("Error updating sheets:", error);
  }
}
