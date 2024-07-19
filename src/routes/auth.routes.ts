import { Router } from "express";
import UserController from "../controllers/users.controller";
import { bodySchemaValidator } from "../middlewares/schema.validator";
import { AddUser, UpdateUser } from "../schemas/user";
import hasPermission from "../middlewares/permission.middleware";
import { CREATE_USER, READ_USER } from "../common/permissions";
import AuthController from "../controllers/auth.controller";
import { LoginData, SignUpData } from "../schemas/auth";

const authRouter = Router({ mergeParams: true });

const authController = new AuthController();

authRouter.post(
  "/signup",
  hasPermission(CREATE_USER),
  bodySchemaValidator(SignUpData),
  authController.signup
);

authRouter.post(
  "/login",
  hasPermission(READ_USER),
  bodySchemaValidator(LoginData),
  authController.login
);

export default authRouter;
