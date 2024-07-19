import { Router } from "express";
import UserController from "../controllers/users.controller";
import { bodySchemaValidator } from "../middlewares/schema.validator";
import { AddUser, UpdateUser } from "../schemas/user";
import hasPermission from "../middlewares/permission.middleware";
import {
  CREATE_USER,
  DELETE_USER,
  READ_USER,
  READ_USERS,
  UPDATE_USER,
} from "../common/permissions";

const userRouter = Router({ mergeParams: true });

const userController = new UserController();

userRouter.get("/", hasPermission(READ_USERS), userController.getAllUsers);
userRouter.post(
  "/",
  hasPermission(CREATE_USER),
  bodySchemaValidator(AddUser),
  userController.createUser
);

userRouter.get(
  "/profile",
  hasPermission(READ_USER),
  userController.getUserProfile
);

userRouter.get("/:id", hasPermission(READ_USER), userController.getUserById);

userRouter.put(
  "/:id",
  hasPermission(UPDATE_USER),
  bodySchemaValidator(UpdateUser),
  userController.updateUser
);

userRouter.delete(
  "/:id",
  hasPermission(DELETE_USER),
  userController.deleteUser
);

export default userRouter;
