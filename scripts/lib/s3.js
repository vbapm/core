import { join } from "node:path";
import fsExtra from "fs-extra";
import dotenv from "dotenv";
import hash from "./hash.js";

const { readFile } = fsExtra;
dotenv.config({ path: join(process.cwd(), ".env") });
const { default: AWS } = await import("aws-sdk");
const Bucket = process.env.VBA_BLOCKS_AWS_S3_BUCKET;

// Load AWS_PROFILE from .env before loading SDK
const s3 = new AWS.S3();

export default {
	async upload({ filename: Key, path }) {
		const Body = await readFile(path);
		const ACL = "public-read";
		const ContentMD5 = hash(Body, { algorithm: "md5", digest: "base64", prefix: false });

		return await s3.upload({ Bucket, Key, Body, ACL, ContentMD5 }).promise();
	}
};

