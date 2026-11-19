/** throwaway: build /tmp/revdemo.docx + /tmp/revdemo-v2.docx for manual review-mode verification */
import { writeFileSync } from 'node:fs'
import { buildDocx } from '../tests/helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

const COMMENTS_XML =
  XML_DECL +
  '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:comment w:id="1" w:author="Alice" w:initials="A" w:date="2026-07-01T10:00:00Z">' +
  '<w:p><w:r><w:t>这个数据需要再核对一下</w:t></w:r></w:p></w:comment>' +
  '</w:comments>'

const body = [
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>市场调研报告(修订演示)</w:t></w:r></w:p>',
  '<w:p><w:r><w:t xml:space="preserve">本季度销量同比增长 </w:t></w:r>' +
    '<w:commentRangeStart w:id="1"/><w:r><w:t>23%</w:t></w:r><w:commentRangeEnd w:id="1"/>' +
    '<w:r><w:commentReference w:id="1"/></w:r>' +
    '<w:r><w:t xml:space="preserve">,超出预期。</w:t></w:r></w:p>',
  '<w:p><w:r><w:t xml:space="preserve">市场规模预计达到 </w:t></w:r>' +
    '<w:del w:id="12" w:author="Bob" w:date="2026-07-02T08:30:00Z">' +
    '<w:r><w:delText>五千亿</w:delText></w:r></w:del>' +
    '<w:ins w:id="11" w:author="Alice" w:date="2026-07-01T10:00:00Z">' +
    '<w:r><w:t>六千亿</w:t></w:r></w:ins>' +
    '<w:r><w:t xml:space="preserve"> 元,竞争格局趋于集中。</w:t></w:r></w:p>',
  '<w:p><w:r><w:t>结论:建议加大渠道投入,关注头部厂商动态。</w:t></w:r></w:p>',
].join('')

const bodyV2 = [
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>市场调研报告(修订演示)</w:t></w:r></w:p>',
  '<w:p><w:r><w:t>本季度销量同比增长 23%,超出预期。</w:t></w:r></w:p>',
  '<w:p><w:r><w:t>新增:海外市场开始放量,值得关注。</w:t></w:r></w:p>',
  '<w:p><w:r><w:t>结论:建议聚焦高端产品线。</w:t></w:r></w:p>',
].join('')

const main = async () => {
  writeFileSync(
    '/tmp/revdemo.docx',
    await buildDocx({
      bodyXml: body,
      extraParts: [
        {
          path: 'word/comments.xml',
          xml: COMMENTS_XML,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
        },
      ],
    }),
  )
  writeFileSync('/tmp/revdemo-v2.docx', await buildDocx({ bodyXml: bodyV2 }))
  console.log('wrote /tmp/revdemo.docx and /tmp/revdemo-v2.docx')
}

void main()
