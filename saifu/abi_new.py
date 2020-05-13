"""
Author: Zhang Chengsheng, @2018.11.22
release: 2019.04.10
"""

import os,sys
from PIL import Image
from matplotlib import pyplot as plt
import numpy as np


Rscript = r'abi_R.R' #[1]outputdir [2]abi file [3]png name [4]refseq


def img_open(img_file):
    return np.array(Image.open(img_file,'r').convert('RGB'))


def abifile_parse(abi_dir):
    print('开始生成峰图图片......')
    outputdir = os.path.join(abi_dir,'pngs')
    if not os.path.exists(outputdir):
        os.mkdir(outputdir)
    genome_type = abi_dir.split('_')[-1]
    abi_files = [i for i in os.listdir(abi_dir) if i.endswith('.ab1')]
    done_list = []
    for i in abi_files:
        name = i.split('.')[0]
        site_chr = str(i.split('.')[1].split('_')[0]).lower()
        site_site = i.split('.')[1].split('_')[1]
        #if i.split('.')[2].split('_')[-1].lower() == 'hg19':
        #    genome_type = 'hg19'
        genome_type = i.split('.')[2].split('_')[-1].lower()
        refseq = refseq_extract(site_chr,site_site,genome_type)
        flank = 10  # same to refseq_extract() flank
        cmd = "rscript {} {} {} {} {} {} {}".format(Rscript,outputdir,os.path.join(abi_dir,i),name,refseq,trans(refseq,reverse=True),flank+1)
        os.system(cmd)
    return outputdir


def refseq_extract(site_chr,site_site,genome_type='hg19'):
    flank = 10
    site_chr = 'chrX' if 'x' in site_chr else site_chr
    if genome_type == 'hg19':
        faidx = r'hg19_1.fa.fai'
        fa = r'hg19_1.fa'
    else:
        fa = r'GRCh38_latest_genomic.fna.filter.fa'
        faidx = r'GRCh38_latest_genomic.fna.filter.fa.fai'
    faidx_dict = {}
    with open(faidx,'r') as o:
        for i in o.readlines():
            faidx_dict[i.strip().split('\t')[0]] = i.strip().split('\t')[1:]
    fbuffer = open(fa,'r')
    start,end = int(site_site)-flank,int(site_site)+flank
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
             "g":"c"}
    new_seq = ""
    if transform:
        for i in seq:
            new_seq += dict1[i]
    else:
        new_seq = seq
    if reverse:
        new_seq = new_seq[::-1]
    return new_seq


def png_reshape(png):
    site_num = 11
    img = img_open(png)
    img1 = img.copy()
    seque = png.split('.')[-2]
    ref = seque.split('-')[-1]
    primary = seque.split('-')[0][0]
    secondary = seque.split('-')[0][-1]
    header_png = '_'.join(png.split('_')[:-1]) + '.header.png'
    line = img1[13, :, :]
    line = np.sum(line, axis=1)
    line_index = np.argwhere(line != 255 * 3).reshape([np.argwhere(line != 255 * 3).shape[0]])
    count = 0
    shadow_idx = []
    for idx, i in enumerate(line_index):
        if idx > 0:
            gap = i - line_index[idx - 1]
            if gap > 10:
                count += 1
                if count == site_num:
                    break
                else:
                    shadow_idx = []
            else:
                shadow_idx.append(i)

    img1_r = img1[:,shadow_idx[0]-20:shadow_idx[0]+30,:]

    img1_1 = np.sum(img1_r,axis=2)
    idx = np.argwhere(img1_1 == 255*3)
    img1_r[idx[:,0],idx[:,1],:] = [230,230,230]

    img2 = img1[0:int(250/300*360),72:1160,:]

    img2[(img2[:, :, 0] < 255) * (img2[:, :, 1] < 255) * (img2[:, :, 2] == 255)] = [0, 0, 250]
    img2[(img2[:, :, 0] < 255) * (img2[:, :, 1] == 255) * (img2[:, :, 2] < 255)] = [0, 200, 0]
    img2[(img2[:, :, 0] == 255) * (img2[:, :, 1] < 255) * (img2[:, :, 2] < 255)] = [230, 0, 0]
    img2[img2.shape[0] - 1 - 8, 28:img2.shape[1] - 1 - 28, :] = [0, 0, 0]
    if ref == primary and not os.path.exists(header_png):
        header = img2[0:30, :, :]
        header = Image.fromarray(np.uint8(header))
        header.save(header_png)
    elif ref == secondary and not os.path.exists(header_png):
        header = img2[30:60, :, :]
        header = Image.fromarray(np.uint8(header))
        header.save(header_png)
    if secondary != primary:
        img2[30:60, :shadow_idx[0]-20-72, :] = [255, 255, 255]
        img2[30:60, shadow_idx[0]+30-72:, :] = [255, 255, 255]
    else:
        img2[30:60, :, :] = [255, 255, 255]
        img2[30:60,shadow_idx[0]-20-72:shadow_idx[0]+30-72:,:] = [230,230,230]

    img2 = Image.fromarray(np.uint8(img2))
    img2.save(png.replace('.png','.mdf.png'))


def png_reshape_batch(dir):
    for i in os.listdir(dir):
        if i.endswith('.png'):
            png_reshape(os.path.join(dir,i))
            #os.remove(os.path.join(dir,i))
