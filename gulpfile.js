var gulp = require('gulp');
var ts = require('gulp-typescript');
var tslint = require('gulp-tslint');
var del = require('del');
var server = require('gulp-develop-server');
var sourcemaps = require('gulp-sourcemaps');
var zip = require('gulp-zip');
var rename = require('gulp-rename');
var jsonTransform = require('gulp-json-transform');
var path = require('path');
var minimist = require('minimist');
var fs = require('fs');

var knownOptions = {
	string: 'packageName',
	string: 'packagePath',
	string: 'specFilter',
	default: {packageName: 'Package.zip', packagePath: path.join(__dirname, '_package'), specFilter: '*'}
};
var options = minimist(process.argv.slice(2), knownOptions);

var tsProject = ts.createProject('./tsconfig.json', {
    // Point to the specific typescript package we pull in, not a machine-installed one
    typescript: require('typescript'),
});

var filesToWatch = ['**/*.ts', '!node_modules/**'];
var filesToLint = ['**/*.ts', '!src/typings/**', '!node_modules/**'];
var staticFiles = ['src/**/*.json', 'src/**/*.csx'];

/**
 * Clean build output.
 */
gulp.task('clean', function () {
    return del([
        'build/**/*',
        // Azure doesn't like it when we delete build/src
        '!build/src'
    ]);
});

/**
 * Generate constants from string tables.
 */
gulp.task('locale:generate', function() {
    gulp.src('./src/locale/en/index.json')
    .pipe(jsonTransform(function(data, file) {
        let result = 
        '/* -------------------------------------------------------------------- */\n' + 
        '/* DO NOT MODIFY, THIS FILE IS GENERATED from src\\locale\\en\\index.json. */\n' + 
        '/* CHECK-IN THIS FILE SO WE CAN TRACK HISTORY.                          */\n' + 
        '/* -------------------------------------------------------------------- */\n' + 
        '\n' + 
        '// tslint:disable-next-line:variable-name\n' + 
        'export const Strings = {\n';
        for (var p in data) {
            if( data.hasOwnProperty(p) ) {
                result += '   "' + p + '": "'+ p + '",\n';
            }
        }
        result += '};\n';
        return result;
    }))
    .pipe(rename('locale.ts'))
    .pipe(gulp.dest('./src/locale'));
});

/**
 * Lint all TypeScript files.
 */
gulp.task('ts:lint', ['locale:generate'], function () {
    return gulp
        .src(filesToLint)
        .pipe(tslint({
            formatter: 'verbose'
        }))
        .pipe(tslint.report({
            summarizeFailureOutput: true
        }));
});

/**
 * Compile TypeScript and include references to library.
 */
gulp.task('ts', ['clean', 'locale:generate'], function() {
    return tsProject
        .src()
        .pipe(sourcemaps.init())
        .pipe(tsProject())
        .pipe(sourcemaps.write('.', {includeContent: false, sourceRoot: '.'}))
        .pipe(gulp.dest('build'));
});

/**
 * Copy statics to build directory.
 */
gulp.task('statics:copy', ['clean'], function () {
    return gulp.src(staticFiles, { base: '.' })
        .pipe(gulp.dest('./build'));
});

/**
 * Build application.
 */
gulp.task('build', ['clean', 'ts:lint', 'ts', 'statics:copy']);

/**
 * Package up app into a ZIP file for Azure deployment.
 */
gulp.task('package', ['build'], function () {
    var packagePaths = [
        'build/**/*',
        'public/**/*',
        'web.config',
        'package.json',
        '**/node_modules/**',
        '!build/src/**/*.js.map', 
        '!build/test/**/*', 
        '!build/test', 
        '!build/src/typings/**/*'];

    //add exclusion patterns for all dev dependencies
    var packageJSON = JSON.parse(fs.readFileSync(path.join(__dirname, 'package.json'), 'utf8'));
    var devDeps = packageJSON.devDependencies;
    for (var propName in devDeps) {
        var excludePattern1 = '!**/node_modules/' + propName + '/**';
        var excludePattern2 = '!**/node_modules/' + propName;
        packagePaths.push(excludePattern1);
        packagePaths.push(excludePattern2);
    }

    return gulp.src(packagePaths, { base: '.' })
        .pipe(zip(options.packageName))
        .pipe(gulp.dest(options.packagePath));
});

gulp.task('server:start', ['build'], function() {
    server.listen({path: 'app.js', cwd: 'build/src'}, function(error) {
        console.error(error);
    });
});

gulp.task('server:restart', ['build'], function() {
    server.restart();
});

gulp.task('default', ['server:start'], function() {
    gulp.watch(filesToWatch, ['server:restart']);
});
