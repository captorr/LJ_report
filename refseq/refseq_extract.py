"""
python3 script
Author: Zhang Chengsheng, @2018.12.28
usage: python script.py [1]
"""

import os,sys,time


def refseq_extract(site_chr,site_site,genome_type='hg19',flank=10):
    if 'hg19' in genome_type:
        faidx = os.path.join(os.path.dirname(sys.argv[0]),'hg19_1.fa.fai')
        fa = os.path.join(os.path.dirname(sys.argv[0]),'hg19_1.fa')
    else:
        fa = os.path.join(os.path.dirname(sys.argv[0]),'GRCh38_latest_genomic.fna.filter.fa')
        faidx = os.path.join(os.path.dirname(sys.argv[0]),'GRCh38_latest_genomic.fna.filter.fa.fai')
    faidx_dict = {}
    with open(faidx,'r') as o:
        for i in o.readlines():
            faidx_dict[i.strip().split('\t')[0]] = i.strip().split('\t')[1:]
    fbuffer = open(fa,'r')
    start,end = int(site_site[0])-flank,int(site_site[1])+flank
    offset = int(faidx_dict[site_chr][1])
    line = int(faidx_dict[site_chr][2])
    size = int(faidx_dict[site_chr][3])
    location = offset + int(start / line) + start - 1
    length = (int(end / line) - int(start / line)) * (size - line) + end - start + 1
    fbuffer.seek(location, 0)
    sequence = fbuffer.read(length).replace('\n','')
    return sequence


def trans(seq, transform=True,reverse=False):
    dict1 = {"A":"T",
             "T":"A",
             "G":"C",
             "C":"G",
             "a":"t",
             "t":"a",
             "c":"g",
             "g":"c",
             "N":"N",
             "n":"n"}
    new_seq = ""
    if transform:
        for i in seq:
            new_seq += dict1[i]
    else:
        new_seq = seq
    if reverse:
        new_seq = new_seq[::-1]
    return new_seq


def main(config):
    out = config.split('.')[0]
    o = open(out + '.fasta','w')
    with open(config,'r') as f:
        for idx,i in enumerate(f.readlines()):
            if idx == 0:
                continue
            lines = i.strip().split('\t')
            chr = lines[2]
            site = sorted([int(lines[3]),int(lines[4])])
            genome = lines[0]
            flank = 600
            seq = refseq_extract(chr,site,genome,flank)
            seq = "{}<{}>{}".format(seq[:flank],seq[flank],seq[flank+1:])
            o.write(">{}_{}_{}\n".format(lines[1],chr,site[0]))
            o.write(seq)
            o.write('\n')
    o.close()


def help():
    sys.stdout.write("\n=======================================================\n")
    sys.stdout.write("===  Author: Zhang Chengsheng                       ===\n")
    sys.stdout.write("===  Usage: No usage                                ===\n")
    sys.stdout.write("=======================================================\n\n")


if __name__ == '__main__':
    help()
    while 1:
        config = input("请输入配置文件:\n")
        if not os.path.isfile(config):
            sys.stdout.write("文件不存在，")
        else:
            break
    try:
        main(config)
        sys.stdout.write('finished!\n')
    except Exception as e:
        sys.stdout.write(e)
        sys.stdout.write('\n')
    finally:
        time.sleep(1000)

#D:\zcs-genex\SCRIPTS\Scripts_for_work\lijuan_refseq\格式.txt