var gulp = require('gulp');
var watch = require('gulp-watch');
var cached = require('gulp-cached');
var jshint = require('gulp-jshint');
var jshintStylish = require('jshint-stylish');
var jscs = require('gulp-jscs');
var jscsStylish = require('jscs-stylish');
var jasmine = require('gulp-jasmine');

var LIB = 'lib/*.js';
var TEST = 'test/*.js';
var EXAMPLES = 'examples/*/*.js';
var SRC = [LIB, TEST, EXAMPLES];

gulp.task('jshint', function () { // linting
	return gulp
		.src(SRC)
		.pipe(cached('jshint-linting'))
		.pipe(jshint({ lookup: true }))
		//.pipe(jshint.reporter('default'))
		.pipe(jshint.reporter(jshintStylish))
		.pipe(jshint.reporter('fail'))
		;
});

gulp.task('jscs', function () { // linting
	return gulp
		.src(SRC)
		.pipe(cached('jscs-linting'))
		.pipe(jscs())
        .pipe(jscs.reporter(jscsStylish))
        .pipe(jscs.reporter('fail'))
        ;
});

gulp.task('test', ['jshint', 'jscs'], function () {
	return gulp
		.src(TEST)
		.pipe(cached('testing'))
		.pipe(jasmine())
		;
});

gulp.task('watch', function () {
	gulp.watch(SRC, ['test']);
});

gulp.task('default', ['watch']);
