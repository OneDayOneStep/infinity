import svelte from 'rollup-plugin-svelte';
import resolve from '@rollup/plugin-node-resolve';
import commonjs from '@rollup/plugin-commonjs';
import css from 'rollup-plugin-css-only';
import sveltePreprocess from "svelte-preprocess";

export default {
	input: './main.svelte',
	output: {
		file: './build/bundle.js',
		format: 'iife',
		name: "infinityMT"
	},
	plugins: [
		svelte({
			preprocess: sveltePreprocess()
		}),
		css({ output: 'bundle.css' }),
		resolve({
			browser: true,
			dedupe: ['svelte']
		}),
		commonjs()
	]
};