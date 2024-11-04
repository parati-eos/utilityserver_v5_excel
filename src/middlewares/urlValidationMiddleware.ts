import { Request, Response, NextFunction } from 'express';
import validator from 'validator';

export const validateUrl = (req: Request, res: Response, next: NextFunction): void => {
  const { url } = req.body;

  if (!url || typeof url !== 'string' || !validator.isURL(url)) {
    res.status(400).json({
      success: false,
      message: 'A valid URL is required in the request body.',
    });
    return;
  }
  next();
};
