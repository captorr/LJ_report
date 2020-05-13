###########################################
##
##  Author: Zhang Chengsheng, 2018.11.23
##	args:[1]outputdir [2]abi file [3]png name [4]refseq1 [5]refseq2(reverse) [6]flank
##
###########################################

suppressPackageStartupMessages(library(sangerseqR, quietly=TRUE))
args<-commandArgs(T)

setwd(args[1])
abi=args[2]
outpng=args[3]
ref=args[4]
ref1=args[5]
flank=as.numeric(args[6])
mismatch=4
pic_width=1200
pic_height=360

hetsangerseq=readsangerseq(abi)
hetcalls=makeBaseCalls(hetsangerseq,ratio=0.2)
CIGAR=c("A","T","C","G")
a=matchPattern(ref,primarySeq(hetcalls),min.mismatch=0,max.mismatch=mismatch)
b=matchPattern(ref1,primarySeq(hetcalls),min.mismatch=0,max.mismatch=mismatch)
##print(a)
##print(b)
if (length(a@ranges@start) == 0 && length(b@ranges@start) == 0){
print(abi)
}
if (length(a@ranges@start) > 0){
trim5=a@ranges@start-1
trim3=primarySeq(hetcalls)@length-a@ranges@start[1]-a@ranges@width[1]+1
f1 = subseq(primarySeq(hetcalls),a@ranges@start,a@ranges@start+a@ranges@width-1)
f2 = subseq(secondarySeq(hetcalls),a@ranges@start,a@ranges@start+a@ranges@width-1)
refseq=substring(ref,flank,flank)
if (!paste0(f2[flank]) %in% CIGAR){
hetcalls=makeBaseCalls(hetsangerseq,ratio=0.5)
a=matchPattern(ref,primarySeq(hetcalls),min.mismatch=0,max.mismatch=mismatch)
f1 = subseq(primarySeq(hetcalls),a@ranges@start,a@ranges@start+a@ranges@width-1)
f2 = subseq(secondarySeq(hetcalls),a@ranges@start,a@ranges@start+a@ranges@width-1)
}
png(paste0(outpng,'.',f1[flank],f2[flank],'-',refseq,'.png'),width=pic_width,height=pic_height,res=150)
chromatogram(hetcalls,ylim=3,width=a@ranges@width[1],height=NA,trim5=trim5,trim3=trim3,showcalls="both",showhets=F)
}else{
if (length(b@ranges@start) > 0){
trim5=b@ranges@start-1
trim3=primarySeq(hetcalls)@length-b@ranges@start[1]-b@ranges@width[1]+1
f1 = subseq(primarySeq(hetcalls),b@ranges@start,b@ranges@start+b@ranges@width-1)
f2 = subseq(secondarySeq(hetcalls),b@ranges@start,b@ranges@start+b@ranges@width-1)
refseq=substring(ref1,flank,flank)
if (!paste0(f2[flank]) %in% CIGAR){
hetcalls=makeBaseCalls(hetsangerseq,ratio=0.5)
b=matchPattern(ref1,primarySeq(hetcalls),min.mismatch=0,max.mismatch=mismatch)
f1 = subseq(primarySeq(hetcalls),b@ranges@start,b@ranges@start+b@ranges@width-1)
f2 = subseq(secondarySeq(hetcalls),b@ranges@start,b@ranges@start+b@ranges@width-1)
}
png(paste0(outpng,'.',f1[flank],f2[flank],'-',refseq,'.png'),width=pic_width,height=pic_height,res=150)
chromatogram(hetcalls,ylim=3,width=b@ranges@width[1],height=NA,trim5=trim5,trim3=trim3,showcalls="both",showhets=F)
}else{
png(paste0(outpng,'.0'))
}
}
dev.off()

