import { Router, type IRouter } from "express";
import healthRouter from "./health";
import documentAiRouter from "./document-ai";
import appStorageRouter from "./app-storage";

const router: IRouter = Router();

router.use(healthRouter);
router.use(documentAiRouter);
router.use(appStorageRouter);

export default router;
