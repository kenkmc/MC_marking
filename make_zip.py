import zipfile
import os
import sys

src_dir = r"D:\python\MC_marking\dist\CheckMate"
out_zip = r"D:\python\MC_marking\dist\CheckMate_v1.6.1.zip"

count = 0
with zipfile.ZipFile(out_zip, 'w', zipfile.ZIP_DEFLATED, compresslevel=1, allowZip64=True) as zf:
    for root, dirs, files in os.walk(src_dir):
        for file in files:
            fpath = os.path.join(root, file)
            arcname = os.path.relpath(fpath, src_dir)
            zf.write(fpath, arcname)
            count += 1
            if count % 1000 == 0:
                print(f"  {count} files...", flush=True)

size_mb = os.path.getsize(out_zip) / 1024 / 1024
print(f"Done: {count} files, {size_mb:.1f} MB")
