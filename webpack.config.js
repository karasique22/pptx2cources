const path = require("path");
const webpack = require("webpack");

module.exports = (env, argv) => ({
	mode: argv.mode === "production" ? "production" : "development",
	cache: {
		type: "filesystem",
	},
	devtool: argv.mode === "production" ? false : "inline-source-map",
	entry: {
		code: "./src/code.ts",
	},
	module: {
		rules: [
			{
				test: /\.tsx?$/,
				use: {
					loader: "ts-loader",
					options: {
						transpileOnly: true,
					},
				},
				exclude: /node_modules/,
			},
		],
	},
	resolve: {
		extensions: [".ts", ".js"],
		fallback: {
			buffer: require.resolve("buffer"),
			stream: require.resolve("stream-browserify"),
			timers: require.resolve("timers-browserify"),
		},
	},
	output: {
		filename: "[name].js",
		path: path.resolve(__dirname, "dist"),
	},
	optimization: {
		minimize: argv.mode === "production",
	},
	performance: {
		hints: argv.mode === "production" ? "warning" : false,
	},
});
