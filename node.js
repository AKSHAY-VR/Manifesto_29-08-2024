import { JSDOM } from 'jsdom';
import * as d3 from 'd3';

// Create a mock DOM
const dom = new JSDOM(`<!DOCTYPE html><html><body><div id="candidates"></div></body></html>`);
const document = dom.window.document;

// Now use D3.js with this document
const container = d3.select(document.querySelector("#candidates"))
    .append("div")
    .attr("class", "candidate-card")
    .text("This is a candidate card");

// Output the resulting HTML
console.log(document.body.innerHTML);
