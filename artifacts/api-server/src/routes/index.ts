import { Router, type IRouter } from "express";
import healthRouter from "./health";
import documentAiRouter from "./document-ai";
import appStorageRouter from "./app-storage";
import attendanceRouter from "./attendance";

const router: IRouter = Router();

router.use(healthRouter);
router.use(documentAiRouter);
router.use(appStorageRouter);
router.use(attendanceRouter);

export default router;
