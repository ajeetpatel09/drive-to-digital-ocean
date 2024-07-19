import { RequestHandler } from "express";
import { buildResponse } from "../common/utils";
import { AppDataSource } from "../data-source";
import UserService from "../services/users.service";
import logger from "../common/logger";
import {
  ResourceNotFoundError,
  ValidationFailedError,
  errorHandler,
} from "../common/errors";
import { CustomRequest } from "../types/types";

class UserController {
  private _usersService = new UserService();

  getAllUsers: RequestHandler = async (req, res) => {
    try {
      const users = await this._usersService.getAllUsers(req.query);
      return res.status(200).send(buildResponse(users, "", ""));
    } catch (error) {
      errorHandler(res, error);
    }
  };

  getUserById: RequestHandler = async (req, res) => {
    try {
      const user = await this._usersService.getUserById(
        parseInt(req.params["id"])
      );
      return res.status(200).send(buildResponse(user, "", ""));
    } catch (error) {
      errorHandler(res, error);
    }
  };

  getUserProfile: RequestHandler = async (req: CustomRequest, res) => {
    try {
      const user = await this._usersService.getUserById(
        req.context.auth?.id || -1
      );
      return res.status(200).send(buildResponse(user, "", ""));
    } catch (error) {
      errorHandler(res, error);
    }
  };

  createUser: RequestHandler = async (req, res) => {
    try {
      const user = await AppDataSource.transaction(
        async (transactionEntityManager) => {
          const user = await this._usersService.createUser(
            transactionEntityManager,
            req.body
          );

          return user;
        }
      );

      return res.send(buildResponse(user, "User created successfully."));
    } catch (error) {
      errorHandler(res, error);
    }
  };

  updateUser: RequestHandler = async (req: CustomRequest, res) => {
    try {
      await AppDataSource.transaction(async (transactionEntityManager) => {
        await this._usersService.updateUser(
          transactionEntityManager,
          req.body,
          parseInt(req.params["id"])
        );

        return;
      });

      return res
        .status(200)
        .send(buildResponse("", "User updated successfully."));
    } catch (error) {
      errorHandler(res, error);
    }
  };

  deleteUser: RequestHandler = async (req, res) => {
    try {
      await AppDataSource.transaction(async (transactionEntityManager) => {
        await this._usersService.deleteUser(
          transactionEntityManager,
          parseInt(req.params["id"])
        );

        return;
      });

      return res
        .status(200)
        .send(buildResponse("", "User deleted successfully."));
    } catch (error) {
      errorHandler(res, error);
    }
  };
}

export default UserController;
