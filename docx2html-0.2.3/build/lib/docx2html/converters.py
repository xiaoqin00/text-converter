import subprocess


def convert_with_abiword(docx_path, file_path):
    """
    This will convert ``file_path`` to docx and place the converted file at
    ``docx_path``
    """
    subprocess.call(
        [
            'abiword',
            '--to=docx',
            '--to-name',
            docx_path,
            file_path,
        ],
    )
