import { Router } from "express";
import userRouter from "./users.routes";
import authRouter from "./auth.routes";

const router = Router({ mergeParams: true });

router.use("/users", userRouter);
router.use("/users/auth", authRouter);

export default router;
