
echo removing calls.csv
rm calls.csv

cat 511_01.csv >> calls.csv
cat 520_01.csv >> calls.csv
cat 514_01.csv >> calls.csv
echo finished:
echo "   lines  words  bytes"

wc calls.csv

