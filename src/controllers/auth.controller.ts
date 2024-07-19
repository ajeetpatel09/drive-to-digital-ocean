import { RequestHandler } from "express";
import { LoginDataSchema, SignUpDataScehema } from "../schemas/auth";
import AuthService from "../services/auth.service";
import { buildResponse } from "../common/utils";
import { errorHandler } from "../common/errors";
import { AppDataSource } from "../data-source";

class AuthController {
  _authService = new AuthService();

  login: RequestHandler = async (req, res) => {
    try {
      const loginData: LoginDataSchema = req.body;
      const { jwtToken, user } = await this._authService.login(loginData);
      return res
        .status(200)
        .send(buildResponse({ jwtToken, user }, "Login successfully"));
    } catch (error) {
      errorHandler(res, error);
    }
  };

  signup: RequestHandler = async (req, res) => {
    try {
      const signUpData: SignUpDataScehema = req.body;

      const { jwtToken, user } = await AppDataSource.transaction(
        async (transactionEntityManager) => {
          const data = await this._authService.signup(
            transactionEntityManager,
            signUpData
          );

          return data;
        }
      );

      return res
        .status(200)
        .send(buildResponse({ jwtToken, user }, "Signup successfully"));
    } catch (error) {
      errorHandler(res, error);
    }
  };
}

export default AuthController;
