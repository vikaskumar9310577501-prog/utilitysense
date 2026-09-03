import fs from 'fs';
import path from 'path';
import os from 'os';

const home = os.homedir();
const pathsToCheck = [
    path.join(home, 'AppData', 'Local', 'Google', 'Chrome', 'User Data', 'Default', 'Local Storage', 'leveldb'),
    path.join(home, 'AppData', 'Local', 'Microsoft', 'Edge', 'User Data', 'Default', 'Local Storage', 'leveldb'),
    path.join(home, 'AppData', 'Local', 'Google', 'Chrome', 'User Data', 'Default', 'IndexedDB'),
    path.join(home, 'AppData', 'Local', 'Microsoft', 'Edge', 'User Data', 'Default', 'IndexedDB'),
    path.join(home, '.gemini', 'antigravity-browser-profile', 'Default', 'Local Storage', 'leveldb'),
    path.join(home, 'AppData', 'Local', 'Temp')
];

function searchDir(dir) {
    if (!fs.existsSync(dir)) return;
    try {
        const files = fs.readdirSync(dir);
        for (const f of files) {
            const full = path.join(dir, f);
            try {
                const stat = fs.statSync(full);
                if (stat.isDirectory()) {
                    searchDir(full);
                } else if (stat.size < 100 * 1024 * 1024) {
                    const content = fs.readFileSync(full);
                    const str = content.toString('latin1');
                    if (str.includes('1020') && (str.includes('2026-08') || str.includes('6818.8') || str.includes('ROORKIE'))) {
                        console.log('FOUND MATCH IN FILE:', full);
                        const lines = str.split('\n');
                        lines.forEach((line) => {
                            if (line.includes('1020') && (line.includes('2026-08') || line.includes('6818.8'))) {
                                console.log('Line snippet:', line.slice(0, 300));
                            }
                        });
                    }
                }
            } catch (err) {}
        }
    } catch (e) {}
}

pathsToCheck.forEach((p) => {
    console.log('Searching in:', p);
    searchDir(p);
});
console.log('Done scanning.');
