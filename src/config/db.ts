import mongoose from 'mongoose';
import dotenv from 'dotenv';

dotenv.config()
const connectDB = async () => {
  try {
    const MONGO_URI = process.env.MONGO_URI
	console.log(MONGO_URI)
		if (MONGO_URI) {
			const conn = await mongoose.connect(MONGO_URI, {
				serverSelectionTimeoutMS: 40000
			});
			console.log(`\x1b[32m \x1b[1m  MongoDB Connected: ${conn.connection.port}\x1b[0m`);
		}
	} catch (error: any) {
		throw error;
	}
};

export default connectDB;
