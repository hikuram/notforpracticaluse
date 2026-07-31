```shell
docker build -t ppt2word-ninja .
docker run --rm   -v /mnt/c/Windows/Fonts:/usr/share/fonts/truetype/custom_ms_fonts:ro   -v $(pwd):/workspace   ppt2word-ninja   bash -c "fc-cache -f && documentbuilder --update-fonts && python3 ppt2word_onlyoffice.py 211021.pptx"
```
