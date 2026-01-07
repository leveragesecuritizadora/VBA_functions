import os

def update_manifest():
    script_dir = os.path.dirname(__file__)
    vbas_dir = os.path.join(script_dir, 'VBAs')
    manifest_path = os.path.join(script_dir, 'manifest.txt')

    bas_files = []
    for root, _, files in os.walk(vbas_dir):
        for file in files:
            if file.endswith('.bas'):
                full_path = os.path.join(root, file)
                relative_path = os.path.relpath(full_path, script_dir)
                bas_files.append(relative_path.replace('\\', '/')) # Use forward slashes for consistency

    with open(manifest_path, 'w') as f:
        for bas_file in bas_files:
            f.write(bas_file + '\n')

    print(f"Updated '{manifest_path}' with {len(bas_files)} .bas file paths from '{vbas_dir}'.")

if __name__ == '__main__':
    update_manifest()
