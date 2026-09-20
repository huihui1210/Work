import { execSync } from 'node:child_process';
import { cpSync, mkdirSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const REPO_URL = 'https://github.com/huihui1210/Work.git';
const PAGE_URL = 'https://huihui1210.github.io/Work/';
const GIT_NAME = 'huihui';
const GIT_EMAIL = '2415541726@qq.com';

const pubDir = join(tmpdir(), 'work-gh-pages-pub');

console.log('1/4 构建产物...');
execSync('npm run build', { stdio: 'inherit' });

console.log('2/4 准备发布目录...');
rmSync(pubDir, { recursive: true, force: true });
mkdirSync(pubDir, { recursive: true });
for (const entry of readdirSync('dist')) {
  cpSync(join('dist', entry), join(pubDir, entry), { recursive: true });
}
writeFileSync(join(pubDir, '.nojekyll'), '');

const run = (cmd) => execSync(cmd, { cwd: pubDir, stdio: 'inherit' });

console.log('3/4 提交到 gh-pages 分支...');
run('git init -b gh-pages');
run('git add .');
run(`git -c user.name=${GIT_NAME} -c user.email=${GIT_EMAIL} commit -m "publish plugin page"`);
run(`git remote add origin ${REPO_URL}`);

console.log('4/4 推送到 GitHub...');
run('git push -f origin gh-pages');
rmSync(pubDir, { recursive: true, force: true });

console.log(`\n发布完成，约 1 分钟后生效：${PAGE_URL}`);
