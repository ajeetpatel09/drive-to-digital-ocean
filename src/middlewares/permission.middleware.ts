import { NextFunction, Response } from "express";
import {
  Permission,
  adminPermissions,
  userPermissions,
} from "../common/permissions";
import { CustomRequest } from "../types/types";
import { ForbiddenError } from "../common/errors";

function hasPermission(requiredPermission: Permission) {
  return async (_req: CustomRequest, _res: Response, next: NextFunction) => {
    if (userPermissions.includes(requiredPermission)) {
      next();
    } else {
      next(
        new ForbiddenError(
          `User  does not have permission ${requiredPermission}`
        )
      );
    }
  };
}

export default hasPermission;
