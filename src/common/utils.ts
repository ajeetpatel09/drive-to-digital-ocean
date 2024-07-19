import { CustomResponse } from "../interfaces/types";
import { saltRounds } from "./constants";
import * as bcrypt from "bcrypt";

export const buildResponse = (data: any, message: string, error: any = "") => {
  const response: CustomResponse = {
    data,
    message,
    error,
  };

  return response;
};

export const encryptString = async (string: string) => {
  return await bcrypt.hash(string, saltRounds);
};
