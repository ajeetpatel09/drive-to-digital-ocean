import "reflect-metadata";
import { AppDataSource } from "./data-source";
import { Request, Response, NextFunction } from "express";

import * as bodyParser from "body-parser";
import * as cors from "cors";
import * as express from "express";
import * as morgan from "morgan";

import { expressjwt } from "express-jwt";
import logger from "./common/logger";
import * as dotenv from "dotenv";
import errorMiddleware from "./middlewares/error.middleware";
import router from "./routes/router";
import * as swaggerJSDoc from "swagger-jsdoc";
import * as swaggerUi from "swagger-ui-express";
import * as cookieParser from "cookie-parser";
import options from "./common/swaggerOptions";
import { jwtCookieName } from "./common/constants";
import { buildResponse } from "./common/utils";

dotenv.config();

morgan.token("host", function (req: express.Request, _res) {
  return req.hostname;
});

const app = express();

app.use(cookieParser());

const specs = swaggerJSDoc(options);

app.use("/api/v1/api-doc", swaggerUi.serve, swaggerUi.setup(specs));

app.use(
  morgan(
    ":date[web] :method :host :url :status :res[content-length] - :response-time ms"
  )
);
app.use(cors({ credentials: true, origin: true }));
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json({ type: "application/json" }));
// app.use(
//   expressjwt({
//     secret: process.env.JWT_SECRET!,
//     getToken: (req) => {
//       if (req.cookies?.[jwtCookieName]) return req.cookies[jwtCookieName];

//       if (req.headers.authorization?.startsWith("Bearer")) {
//         return req.headers.authorization.split("Bearer ")[1];
//       }

//       return req.headers?.authorization;
//     },
//     algorithms: ["HS256"],
//     requestProperty: "context.auth",
//   }).unless({
//     path: [
//       // mention path where you dont wanna check token in following format RegExp("/api/v1/user/auth"),
//       RegExp("/api/v1/users/auth"),
//     ],
//   })
// );

// app.use((err: any, _req: Request, res: Response, next: NextFunction) => {
//   if (err.name === "UnauthorizedError") {
//     res.status(403).send(buildResponse("", "invalid token", err));
//   } else {
//     next(err);
//   }
// });

const port = process.env.PORT || 80;

app.use("/api/v1", router);

app.use(errorMiddleware);

app.listen(port, async () => {
  logger.info("App Started on port", { port });

  try {
    await AppDataSource.initialize();
    console.log("Database connection successful...");
  } catch (error) {
    logger.error(error);
  }
});
